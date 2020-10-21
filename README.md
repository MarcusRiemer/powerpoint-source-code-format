# Powerpoint Source Code Formatter

Dealing with source code in Powerpoint is cumbersome. Pasting the sourcecode works for certain applications
that fill the clipboard with RTF or HTML data, but keeping all sorts of snippets up to date is rather annoying. 
This plugin adds ribbon UI elements which run [pygments](https://pygments.org/) to apply syntax highlighting 
to selected text-shapes.

![GIF-Video that shows the plugin in action](https://playground.marcusriemer.de/pp_source_code.gif)

# Hacky work-in-progress warning: This may destroy the formatted sources!

I have only tested this Plugin with a few of my slides and during development I stumbled over all sorts of
hacks that I needed to employ to achieve the seemingly stable formatting that you can see in the video.
During development I had to replace spaces with non breaking space (and vice versa), learned about the 
`\v` escape sequence that Powerpoint seems to use internally for line breaks that are not paragraphs,
dug through the annoyances of the [Win32 HTML clipboard format](https://docs.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format)
and possibly some more things.

That being said: <kbd>Ctrl</kbd> + <kbd>Z</kbd> should work if something goes wrong.

# How to install Pygments

Powerpoint Source Code Formatter assumes that a Python installation with an installed `pygmentize.exe` is 
available in the system wide `%PATH%` variable. Pull requests to make this better discoverable are welcome.

A compatible `pygmentize.exe` can be achived with the following these steps:

* Install Python 3 from [python.org](https://www.python.org/) and ensure that the "Add Python to PATH" option
  is checked.
* Run `easy_install Pygments` from the commandline shell, this will actually install the used syntax 
  highlighting program
  
  * If `easy_install.exe` can't be found you probably haven't added Python to your PATH. Rerun the installer
    to ensure that Python is added to PATH.
  * If you have installed Python for all users, you will need to run the shell as admin
* You can run `pygmentize.exe --version` to ensure that the installation was successful.