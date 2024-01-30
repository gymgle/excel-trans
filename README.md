# Excel-trans

`python 3.11` with `Flet`

Etranslate Excel files and keep the original cell format.

## What's it looking like?
![screenshot](./assets/app_light.png)

![screenshot](./assets/app_dark.png)

## How to use?

Download binary from release or clone this repository run with python.

## As a Developer

```shell
# 1. Clone this project.
$ git clone https://github.com/gymgle/excel-trans.git

# 2. Install pip requirements. 
$ cd excel-trans
$ pip3 install -r requirements.txt

# 3. Try it!
$ python main.py
```

## How to Build?

Build Windows/Linux/macOS executable binary file via PyInstaller.

1. Install PyInstaller
    ``` shell
    $ pip install pyinstaller
    ```

2. Prepare UPX (Optional)

    Download UPX [Here](https://github.com/upx/upx/releases), put `upx.exe` (Windows) to project dir.

3. Package project
    ```shell
    $ flet pack main.py
    ```

You can find the packaged `main` in `dist` dir.
