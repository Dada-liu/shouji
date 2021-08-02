# tkinter模块

python的GUI模块



GUI属性有些不能跨平台，因为不同的OS平台提供的库不同

让Toplevel窗口置顶和取消置顶：

```python
window = tk.Tk()
Child_window = Toplevel(window)
Child_window.attributes("-topmost", True)#置顶
#Child_window.wm_attributes("-topmost", True)
Child_window.attributes("-topmost",False)#取消置顶
#Child_window.wm_attributes("-topmost",False)
```

将窗口设置为工具窗口

```python
Child_window.attributes("-toolwindow", True)
```

设置窗口的透明度：

```python
Child_window.attributes('-alpha', 0.6)
```