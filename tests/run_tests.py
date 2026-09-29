import json
import sys
import time
import unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'));sys.path.insert(0,str(ROOT/'tests'))
from common import write_json,write_text,code_hash
import test_acceptance as t
import test_v2 as v2

class Recorded(unittest.TextTestResult):
    def startTest(self,test):self.started=time.monotonic();super().startTest(test)
    def addSuccess(self,test):
        super().addSuccess(test);self.capture(test,'PASS','')
    def addFailure(self,test,err):
        super().addFailure(test,err);self.capture(test,'FAIL',self._exc_info_to_string(err,test))
    def addError(self,test,err):
        super().addError(test,err);self.capture(test,'ERROR',self._exc_info_to_string(err,test))
    def addSkip(self,test,reason):
        super().addSkip(test,reason);self.capture(test,'NOT_EXECUTED',reason)
    def capture(self,test,status,error):
        key=test._testMethodName.split('_')[1]
        self.records.append(dict(id=key,name=test._testMethodName,status=status,elapsed_seconds=round(time.monotonic()-self.started,3),
                                 evidence=t.OBSERVATIONS.get(key,{}),error=error))
    records=[]

suite=unittest.defaultTestLoader.loadTestsFromTestCase(t.Acceptance)
suite.addTests(unittest.defaultTestLoader.loadTestsFromTestCase(v2.Improvements))
r=unittest.TextTestRunner(verbosity=2,resultclass=Recorded).run(suite)
payload=dict(code_fingerprint=code_hash(),executed=r.testsRun,passed=len(r.records)-len(r.errors)-len(r.failures)-len(r.skipped),
             failures=len(r.failures),errors=len(r.errors),not_executed=len(r.skipped),results=r.records,
             boundary='测试仅证明合成输入下对应行为；不证明材料真实或项目可投。受限功能拒绝测试通过≠功能已实现。')
write_json(ROOT/'outputs'/'测试执行结果.json',payload)
lines=['# 实际测试结果','',payload['boundary'],'',f'执行 {payload["executed"]}；通过 {payload["passed"]}；失败 {payload["failures"]}；错误 {payload["errors"]}；未执行 {payload["not_executed"]}。','',
       '|编号|用例|结果|耗时（秒）|','|---|---|---|---:|']
for row in r.records:lines.append(f'|{row["id"]}|{row["name"]}|{row["status"]}|{row["elapsed_seconds"]}|')
for row in r.records:
    if row['error']:lines += ['',f'## {row["id"]} 失败详情','', '```',row['error'],'```']
write_text(ROOT/'outputs'/'测试执行结果.md','\n'.join(lines)+'\n')
sys.exit(0 if r.wasSuccessful() else 1)
