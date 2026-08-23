# DOC integration fixtures

These fixtures are harmless synthetic payloads used only by
`ConverterIntegration.Tests.ps1`. They contain no personal or business content
and are reconstructed into Pester `TestDrive` only; no runtime document is
written to the repository.

| Fixture | Format | SHA-256 | Redistribution status |
|---|---|---|---|
| `special-name.doc.base64` | Word 97-2003 binary DOC | `2a7c5606fb6f7aeed9b88e6a72f162665bc7c5849ba0d475b5b9479a0dabbedf` | Synthetic fixture generated for this repository |
| `password-protected.doc.base64` | Password-protected Word 97-2003 binary DOC | `bb3c76697092123e2b21165b0e2021aa353f355dc23995f4ab2ea8ba0c972435` | Synthetic negative-path fixture generated for this repository |
| `corrupt.doc.base64` | Deliberately invalid legacy-document byte payload | `570e5ac66722e9169fee82ef08036e6eebde2741ba0b91f26fdf3050f7bddcc1` | Synthetic negative fixture safe to redistribute with this repository |

The integration suite uses the binary fixtures to validate the real Word COM
conversion boundary, including a wrong-password failure, while preserving the
source document and leaving no staging artifacts behind.
