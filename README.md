# iis-oracle

Measured classic ASP / VBScript semantics from a real IIS, via GitHub Actions
Windows runners. Exists so that engine-behaviour questions (chiefly for
[AxonASP](https://github.com/guimaraeslucas/axonasp)) are settled by
measurement, never derived from documentation or memory.

Every run prints the answering engine's exact version first, so a measurement
is citable as "IIS 10.0 / VBScript x.y.z, run #N" with a permanent log URL.

## How it works

`probes/*.vbs` run under `cscript` (pure VBScript semantics, no web server).
`probes/*.asp` are served by a real IIS with classic ASP enabled on a
`windows-2022` runner and fetched over HTTP. Results land in the job's step
summary and the log.

Probe pages must be self-contained (numeric ADO constants, no includes, no
database, no outbound calls), must trap errors per-case with
`On Error Resume Next` so one failure doesn't hide the rest, and — for parser
questions — must go one construct per page, because a compile error takes the
whole page down. Runtime questions can share a page.

## Running

Pushes to `main` run everything, or trigger manually:

```sh
gh workflow run measure -R saborrie/iis-oracle && gh run watch -R saborrie/iis-oracle
```

## Caveat

Runners are Windows Server 2022+; the reference production environment this
replaces was Server 2016 (VBScript 5.8.16384). Core VBScript and ADO
semantics are frozen across those builds, but for a question that plausibly
changed between Windows builds, confirm against a matching build before
relying on the answer.
