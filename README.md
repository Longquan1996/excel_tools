# excel_tools

`pyinstaller --additional-hooks-dir=./hooks --clean --onefile --windowed app.py`   
`pyinstaller --hidden-import=openpyxl app.py`   
`pyinstaller --exclude-module numpy --exclude-module scipy app.py`   
`--clean` 删除之前生成的构建缓存、临时文件和输出目录中的文件。这可以确保你生成的是一个干净的构建，没有任何残留的旧文件或缓存。   
`--onefile` 将所有文件和依赖打包成一个单独的 `.exe` 文件。虽然便于分发和使用，但文件较大，启动时会稍慢，因为需要解压临时文件。   
`--onedir` 生成一个包含 .exe 和所有必要文件的文件夹。这样可以避免单个 `.exe` 文件的大小过大，因为所有文件都被分散在目录中。   
`--windowed` 不会显示控制台窗口（即命令行窗口）。这是为了让 GUI 应用程序看起来更加清爽，用户不需要看到额外的控制台输出。   
