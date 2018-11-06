# Core-Python-for-python3.6
### 需要注意python版本对库的影响和库自身的改动 ###
1. Tkinter 在python3.6中为tkinter，其中的messagebox，而不是MessageBox
2. 注意库pywin32 和pypiwin32 的更新，参考：> https://blog.csdn.net/qq_41703291/article/details/80433071.
3. 注意`# xl = win32.gencache.ensuredispath('%s.Application' % app)`	静态调度，需要 COM Makepy utility 为有前提的静态调度，需要改为`win32.Dispatch('%s.Application' % pp)`
4. `tk().withdraw()`在python3.6 中并不存在，需要用tk.Tk()再调用withdraw（）函数，或者`root = tk.Tk()`  `root.withdraw()`

echo “增加readme.md介绍文件”
git init 初始化git
git add file.. 增加文件
git commit -m "描述本次push"
git remote add origin git。。。。 增加git origin，第一次可使用
git push -u origin master 把增加的内容push 上去。

