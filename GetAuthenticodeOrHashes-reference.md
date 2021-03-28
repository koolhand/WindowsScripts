# GetAuthenticodeOrHashes.ps1

Luke Morey 2021-03-28

For a given folder, export a report of...
- authenticode signers where available
- SHA-256 hash of executabl files where Authenticode signature is missing, excluding zero byte files
- list of zero byte files

## TO DO
- distinguish invalid sig (HashMismatch,NotTrusted) that exists from missing sig (NotSigned)
- see if you can make it cleaner with -ExpandProperty
- will it handle all filenames? Do you need LiteralPath or any of the tricks from https://codereview.stackexchange.com/questions/223328/creating-a-file-of-md5-hashes-for-all-files-in-a-directory-in-powershell

## reference
- https://docs.microsoft.com/en-us/windows/win32/seccrypto/cryptography-tools#introduction-to-code-signing
- Get-AuthenticodeSignature https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/get-authenticodesignature
- Get-FileHash https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-filehash
- https://pentestlab.blog/tag/authenticode/ for the limitations / workarounds

```enum System.Management.Automation.SignatureStatus
Name value
---- -----
Valid 0
UnknownError 1
NotSigned 2
HashMismatch 3
NotTrusted 4
NotSupportedFileFormat 5
```
