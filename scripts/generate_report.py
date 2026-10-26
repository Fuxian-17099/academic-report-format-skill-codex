import argparse,json
from pathlib import Path
from engine import generate
def main():
 p=argparse.ArgumentParser(description='Generate Academic Report Format 1.0.0 DOCX from JSON')
 p.add_argument('input',type=Path);p.add_argument('output',type=Path);a=p.parse_args()
 try:generate(json.loads(a.input.read_text(encoding='utf-8-sig')),a.output,a.input.resolve().parent)
 except (ValueError,KeyError,FileNotFoundError) as e:p.exit(2,'ERROR: '+str(e)+'\n')
 print('PASS generated '+str(a.output.resolve())+'; update fields in Word before delivery')
if __name__=='__main__':main()
