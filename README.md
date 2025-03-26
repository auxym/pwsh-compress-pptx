# pwsh-compress-pptx

Powershell module providing tools to make Powerpoint .pptx file smaller. This
module currently exports a single cmdlet: `Compress-PptxMedia`. This cmdlet
uses ffmpeg to re-encode videos with the AV1 codec.

Dependencies: the `ffmpeg` and `ffprobe` must be on your path.
[Scoop](https://scoop.sh/) is an easy way to install ffmpeg on Windows.
