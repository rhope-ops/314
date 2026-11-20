# Ion round 3: more file readers (plan, approved by the PAM 2026-11-20 in the question dialog)

Start after PR #14 (round 2) is merged. Each reader is its own module (src/<Name>.psm1, listed in
scripts/build_paste.py MODULES on the PAM's request of 2026-11-20), .NET only, with tests on invented data.
Done = all built, `scripts/check.sh` CHECK PASS, fresh review without MAJOR, merged, bundle rebuilt.

1. Zip archives (.zip): files inside read in memory with the existing readers; limits on depth (1), entry count
   and unpacked size (zip-bomb guard); each inner file reported as "archive.zip/inner/path".
2. Small readers: .rtf (control words stripped), .eml (headers and text body; base64/quoted-printable parts),
   .odt/.ods/.odp (zip + content.xml), .vsdx (zip + visio/pages/*.xml).
3. Legacy .doc/.xls/.ppt and Outlook .msg (OLE): text runs (UTF-16LE and ANSI) pulled from the binary; a reason
   when none is found. Partial by design: say so in the README.
4. Images (.jpg/.jpeg/.png/.tif/.tiff/.bmp/.gif) and scanned PDFs (DCTDecode JPEG images in an image-only PDF):
   Windows OCR (Windows.Media.Ocr through WinRT, no module). First the PAM runs a test snippet on the gov device;
   without OCR there (error or no language), images stay NOT SCANNED with a reason. A time cap per file.
