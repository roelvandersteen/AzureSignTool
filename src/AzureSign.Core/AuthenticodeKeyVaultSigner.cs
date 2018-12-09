﻿using AzureSign.Core.Interop;
using Microsoft.Extensions.Logging;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace AzureSign.Core
{
    /// <summary>
    /// Signs a file with an Authenticode signature.
    /// </summary>
    public class AuthenticodeKeyVaultSigner : IDisposable
    {
        private readonly AsymmetricAlgorithm _signingAlgorithm;
        private readonly X509Certificate2 _signingCertificate;
        private readonly HashAlgorithmName _fileDigestAlgorithm;
        private readonly TimeStampConfiguration _timeStampConfiguration;
        private readonly MemoryCertificateStore _certificateStore;
        private readonly X509Chain _chain;
        private readonly SignCallback _signCallback;
        private static string ManifestLocation;

        /// <summary>
        /// Creates a new instance of <see cref="AuthenticodeKeyVaultSigner" />.
        /// </summary>
        /// <param name="signingAlgorithm">
        /// An instance of an asymmetric algorithm that will be used to sign. It must support signing with
        /// a private key.
        /// </param>
        /// <param name="signingCertificate">The X509 public certificate for the <paramref name="signingAlgorithm"/>.</param>
        /// <param name="fileDigestAlgorithm">The digest algorithm to sign the file.</param>
        /// <param name="timeStampConfiguration">The timestamp configuration for timestamping the file. To omit timestamping,
        /// use <see cref="TimeStampConfiguration.None"/>.</param>
        /// <param name="additionalCertificates">Any additional certificates to assist in building a certificate chain.</param>
        public AuthenticodeKeyVaultSigner(AsymmetricAlgorithm signingAlgorithm, X509Certificate2 signingCertificate,
            HashAlgorithmName fileDigestAlgorithm, TimeStampConfiguration timeStampConfiguration,
            X509Certificate2Collection additionalCertificates = null)
        {
            _fileDigestAlgorithm = fileDigestAlgorithm;
            _signingCertificate = signingCertificate ?? throw new ArgumentNullException(nameof(signingCertificate));
            _timeStampConfiguration = timeStampConfiguration ?? throw new ArgumentNullException(nameof(timeStampConfiguration));
            _signingAlgorithm = signingAlgorithm ?? throw new ArgumentNullException(nameof(signingAlgorithm));
            _certificateStore = MemoryCertificateStore.Create();
            _chain = new X509Chain();
            if (additionalCertificates != null)
            {
                _chain.ChainPolicy.ExtraStore.AddRange(additionalCertificates);
            }
            //We don't care about the trustworthiness of the cert. We just want a chain to sign with.
            _chain.ChainPolicy.VerificationFlags = X509VerificationFlags.AllFlags;


            if (!_chain.Build(signingCertificate))
            {
                throw new InvalidOperationException("Failed to build chain for certificate.");
            }
            for (var i = 0; i < _chain.ChainElements.Count; i++)
            {
                _certificateStore.Add(_chain.ChainElements[i].Certificate);
            }
            _signCallback = SignCallback;

            ThrowIfNotInitialized();
        }

        /// <summary>Authenticode signs a file.</summary>
        /// <param name="pageHashing">True if the signing process should try to include page hashing, otherwise false.
        /// Use <c>null</c> to use the operating system default. Note that page hashing still may be disabled if the
        /// Subject Interface Package does not support page hashing.</param>
        /// <param name="descriptionUrl">A URL describing the signature or the signer.</param>
        /// <param name="description">The description to apply to the signature.</param>
        /// <param name="path">The path to the file to signed.</param>
        /// <param name="logger">An optional logger to capture signing operations.</param>
        /// <returns>A HRESULT indicating the result of the signing operation. S_OK, or zero, is returned if the signing
        /// operation completed successfully.</returns>
        public unsafe int SignFile(ReadOnlySpan<char> path, ReadOnlySpan<char> description, ReadOnlySpan<char> descriptionUrl, bool? pageHashing, ILogger logger = null)
        {
            void CopyAndNullTerminate(ReadOnlySpan<char> str, Span<char> destination)
            {
                str.CopyTo(destination);
                destination[destination.Length - 1] = '\0';
            }

            using (var ctx = new Kernel32.ActivationContext(ManifestLocation))
            {
                var flags = SignerSignEx3Flags.SIGN_CALLBACK_UNDOCUMENTED;
                if (pageHashing == true)
                {
                    flags |= SignerSignEx3Flags.SPC_INC_PE_PAGE_HASHES_FLAG;
                }
                else if (pageHashing == false)
                {
                    flags |= SignerSignEx3Flags.SPC_EXC_PE_PAGE_HASHES_FLAG;
                }

                SignerSignTimeStampFlags timeStampFlags;
                ReadOnlySpan<byte> timestampAlgorithmOid;
                string timestampUrl;
                switch (_timeStampConfiguration.Type)
                {
                    case TimeStampType.Authenticode:
                        timeStampFlags = SignerSignTimeStampFlags.SIGNER_TIMESTAMP_AUTHENTICODE;
                        timestampAlgorithmOid = default;
                        timestampUrl = _timeStampConfiguration.Url;
                        break;
                    case TimeStampType.RFC3161:
                        timeStampFlags = SignerSignTimeStampFlags.SIGNER_TIMESTAMP_RFC3161;
                        timestampAlgorithmOid = AlgorithmTranslator.HashAlgorithmToOidAsciiTerminated(_timeStampConfiguration.DigestAlgorithm);
                        timestampUrl = _timeStampConfiguration.Url;
                        break;
                    default:
                        timeStampFlags = 0;
                        timestampAlgorithmOid = default;
                        timestampUrl = null;
                        break;
                }

                Span<char> pathWithNull = path.Length > 0x100 ? new char[path.Length + 1] : stackalloc char[path.Length + 1];
                Span<char> descriptionBuffer = description.Length > 0x100 ? new char[description.Length + 1] : stackalloc char[description.Length + 1];
                Span<char> descriptionUrlBuffer = descriptionUrl.Length > 0x100 ? new char[descriptionUrl.Length + 1] : stackalloc char[descriptionUrl.Length + 1];
                Span<char> timestampUrlBuffer = timestampUrl == null ?
                    default : timestampUrl.Length > 0x100 ?
                    new char[timestampUrl.Length + 1] : stackalloc char[timestampUrl.Length + 1];

                CopyAndNullTerminate(path, pathWithNull);
                CopyAndNullTerminate(description, descriptionBuffer);
                CopyAndNullTerminate(descriptionUrl, descriptionUrlBuffer);
                if (timestampUrl != null)
                {
                    CopyAndNullTerminate(timestampUrl, timestampUrlBuffer);
                }

                fixed (byte* pTimestampAlgorithm = timestampAlgorithmOid)
                fixed (char* pTimestampUrl = timestampUrlBuffer)
                fixed (char* pPath = pathWithNull)
                fixed (char* pDescription = descriptionBuffer)
                fixed (char* pDescriptionUrl = descriptionUrlBuffer)
                {
                    var fileInfo = new SIGNER_FILE_INFO(pPath, default);
                    var subjectIndex = 0u;
                    var signerSubjectInfoUnion = new SIGNER_SUBJECT_INFO_UNION(&fileInfo);
                    var subjectInfo = new SIGNER_SUBJECT_INFO(&subjectIndex, SignerSubjectInfoUnionChoice.SIGNER_SUBJECT_FILE, signerSubjectInfoUnion);
                    var authCodeStructure = new SIGNER_ATTR_AUTHCODE(pDescription, pDescriptionUrl);
                    var storeInfo = new SIGNER_CERT_STORE_INFO(
                        dwCertPolicy: SignerCertStoreInfoFlags.SIGNER_CERT_POLICY_CHAIN,
                        hCertStore: _certificateStore.Handle,
                        pSigningCert: _signingCertificate.Handle
                    );
                    var signerCert = new SIGNER_CERT(
                        dwCertChoice: SignerCertChoice.SIGNER_CERT_STORE,
                        union: new SIGNER_CERT_UNION(&storeInfo)
                    );
                    var signatureInfo = new SIGNER_SIGNATURE_INFO(
                        algidHash: AlgorithmTranslator.HashAlgorithmToAlgId(_fileDigestAlgorithm),
                        psAuthenticated: IntPtr.Zero,
                        psUnauthenticated: IntPtr.Zero,
                        dwAttrChoice: SignerSignatureInfoAttrChoice.SIGNER_AUTHCODE_ATTR,
                        attrAuthUnion: new SIGNER_SIGNATURE_INFO_UNION(&authCodeStructure)
                    );
                    var callbackPtr = Marshal.GetFunctionPointerForDelegate(_signCallback);
                    var signCallbackInfo = new SIGN_INFO(callbackPtr);

                    logger?.LogTrace("Getting SIP Data");
                    var sipKind = SipExtensionFactory.GetSipKind(path);
                    void* sipData = (void*)0;
                    IntPtr context = IntPtr.Zero;

                    switch (sipKind)
                    {
                        case SipKind.Appx:
                            APPX_SIP_CLIENT_DATA clientData;
                            SIGNER_SIGN_EX3_PARAMS parameters;
                            clientData.pSignerParams = &parameters;
                            sipData = &clientData;
                            flags &= ~SignerSignEx3Flags.SPC_INC_PE_PAGE_HASHES_FLAG;
                            flags |= SignerSignEx3Flags.SPC_EXC_PE_PAGE_HASHES_FLAG;
                            FillAppxExtension(ref clientData, flags, timeStampFlags, &subjectInfo, &signerCert, &signatureInfo, &context, pTimestampUrl, pTimestampAlgorithm, &signCallbackInfo);
                            break;
                    }

                    logger?.LogTrace("Calling SignerSignEx3");
                    var result = mssign32.SignerSignEx3
                    (
                        flags,
                        &subjectInfo,
                        &signerCert,
                        &signatureInfo,
                        IntPtr.Zero,
                        timeStampFlags,
                        pTimestampAlgorithm,
                        pTimestampUrl,
                        IntPtr.Zero,
                        sipData,
                        &context,
                        IntPtr.Zero,
                        &signCallbackInfo,
                        IntPtr.Zero
                    );
                    if (result == 0 && context != IntPtr.Zero)
                    {
                        Debug.Assert(mssign32.SignerFreeSignerContext(context) == 0);
                    }
                    if (result == 0 && sipKind == SipKind.Appx)
                    {
                        var state = ((APPX_SIP_CLIENT_DATA*)sipData)->pAppxSipState;
                        if (state != IntPtr.Zero)
                        {
                            Marshal.Release(state);
                        }
                    }
                    return result;
                }
            }            
        }

        /// <summary>
        /// Frees all resources used by the <see cref="AuthenticodeKeyVaultSigner" />.
        /// </summary>
        public void Dispose()
        {
            _chain.Dispose();
            _certificateStore.Close();
        }

        private unsafe int SignCallback(
            IntPtr pCertContext,
            IntPtr pvExtra,
            uint algId,
            byte[] pDigestToSign,
            uint dwDigestToSign,
            ref CRYPTOAPI_BLOB blob
        )
        {
            const int E_INVALIDARG = unchecked((int)0x80070057);
            byte[] digest;
            switch (_signingAlgorithm)
            {
                case RSA rsa:
                    digest = rsa.SignHash(pDigestToSign, _fileDigestAlgorithm, RSASignaturePadding.Pkcs1);
                    break;
                case ECDsa ecdsa:
                    digest = ecdsa.SignHash(pDigestToSign);
                    break;
                default:
                    return E_INVALIDARG;
            }
            var resultPtr = Marshal.AllocHGlobal(digest.Length);
            Marshal.Copy(digest, 0, resultPtr, digest.Length);
            blob.pbData = resultPtr;
            blob.cbData = (uint)digest.Length;
            return 0;
        }

        private static unsafe void FillAppxExtension(
            ref APPX_SIP_CLIENT_DATA clientData,
            SignerSignEx3Flags flags,
            SignerSignTimeStampFlags timestampFlags,
            SIGNER_SUBJECT_INFO* signerSubjectInfo,
            SIGNER_CERT* signerCert,
            SIGNER_SIGNATURE_INFO* signatureInfo,
            IntPtr* signerContext,
            char* timestampUrl,
            byte* timestampOid,
            SIGN_INFO* signInfo
        )
        {
            clientData.pSignerParams->dwFlags = flags;
            clientData.pSignerParams->dwTimestampFlags = timestampFlags;
            clientData.pSignerParams->pSubjectInfo = signerSubjectInfo;
            clientData.pSignerParams->pSignerCert = signerCert;
            clientData.pSignerParams->pSignatureInfo = signatureInfo;
            clientData.pSignerParams->ppSignerContext = signerContext;
            clientData.pSignerParams->pwszHttpTimeStamp = timestampUrl;
            clientData.pSignerParams->pszTimestampAlgorithmOid = timestampOid;
            clientData.pSignerParams->pSignCallBack = signInfo;

        }

        /// <summary>
        /// It is required to call initialize before using the AzureSign to ensure dll's are loaded correctly.
        /// The initialize should happen as early as possible in the host process.
        /// </summary>
        /// <param name="assemblyPath">Override the default location for the assembly. Necessary when running on Azure App Services</param>
        public static void Initialize(string assemblyPath = null)
        {
            if (ManifestLocation != null)
                return; // already initialized


            var is64bit = IntPtr.Size == 8;

            if(string.IsNullOrWhiteSpace(assemblyPath))
            {
                // the directory should be next to the assembly where this type is.
                assemblyPath = Path.GetDirectoryName(typeof(AuthenticodeKeyVaultSigner).Assembly.Location);
            }            

            var basePath = Path.Combine(assemblyPath, is64bit ? "x64" : "x86");
            ManifestLocation = Path.Combine(assemblyPath, is64bit ? "x64" : "x86", "SignTool.exe.manifest");

            //
            // Ensure we invoke wintrust!DllMain before we get too far.
            // This will call wintrust!RegisterSipsFromIniFile and read in wintrust.dll.ini
            // to swap out some local SIPs. Internally, wintrust will call LoadLibraryW
            // on each DLL= entry, so we need to also adjust our DLL search path or we'll
            // load unwanted system-provided copies.
            //
            Kernel32.SetDllDirectoryW(basePath);
            Kernel32.LoadLibraryW($@"{basePath}\wintrust.dll");
            Kernel32.LoadLibraryW($@"{basePath}\mssign32.dll");
        }

        private static void ThrowIfNotInitialized()
        {
            if(ManifestLocation == null)
            {
                throw new InvalidOperationException("Initialize must be called early in the host process");
            }
        }
    }   
}
