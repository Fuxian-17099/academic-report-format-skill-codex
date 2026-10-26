"""Use a separate native Word instance, update fields, save and optionally export PDF."""
import argparse,subprocess,json
from pathlib import Path
def update(path,pdf=None):
 path=Path(path).resolve()
 if not path.is_file():raise FileNotFoundError(path)
 # JSON stdin is data, never interpolated PowerShell code.
 code=r'''
$ErrorActionPreference='Stop'
$cfg=[Console]::In.ReadToEnd() | ConvertFrom-Json
$word=$null; $doc=$null
try {
 $word=New-Object -ComObject Word.Application
 $word.Visible=$false; $word.DisplayAlerts=0; $word.AutomationSecurity=3
 $doc=$word.Documents.Open($cfg.path,$false,$false,$false)
 for($pass=0;$pass -lt 3;$pass++) {
  $null=$doc.Fields.Update()
  foreach($story in $doc.StoryRanges) { $range=$story; while($null -ne $range) { $null=$range.Fields.Update(); $range=$range.NextStoryRange } }
  foreach($toc in $doc.TablesOfContents) { $toc.Update() }
  $doc.Repaginate()
 }
 $doc.Save()
 $pages=$doc.ComputeStatistics(2)
 if($cfg.pdf) {$doc.ExportAsFixedFormat($cfg.pdf,17)}
 Write-Output ('PASS Word fields updated; pages='+$pages)
} finally {
 if($null -ne $doc) {$doc.Close(0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)}
 if($null -ne $word) {$word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}
}
'''
 args={'path':str(path),'pdf':str(Path(pdf).resolve()) if pdf else None}
 r=subprocess.run(['powershell.exe','-NoProfile','-NonInteractive','-Command',code],input=json.dumps(args),capture_output=True,text=True,timeout=180)
 if r.returncode:raise RuntimeError(r.stderr.strip() or r.stdout.strip())
 from package_utils import finalize
 finalize(path)
 return r.stdout.strip()
def main():
 p=argparse.ArgumentParser();p.add_argument('input');p.add_argument('--pdf');a=p.parse_args()
 try:print(update(a.input,a.pdf))
 except Exception as e:p.exit(2,'FAIL Word refresh unavailable: '+str(e)+'\nKeep updateFields=true and use Ctrl+A, F9 in Word, then update the entire TOC.\n')
if __name__=='__main__':main()
