import sys
sys.path.insert(0, r'C:/Users/xy24/.claude/skills/ppt-acceptance-check')
from layers import runs
import inspect
src = inspect.getsource(runs._walk_runs)
print('runs.py import OK')
print('has italic capture:', 'italic = int' in src)
print('dims default has italic:', 'italic' in inspect.getsource(runs.run))
