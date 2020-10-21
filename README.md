# Powerpoint Source Code Formatter

Dealing with source code in Powerpoint is cumbersome. Pasting the sourcecode works for certain applications
that fill the clipboard with RTF or HTML data, but keeping all sorts of snippets up to date is cumbersome. 
This plugin adds ribbon UI elements which run [pygments](https://pygments.org/) to apply syntax highlighting 
to selected texts or shapes.



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