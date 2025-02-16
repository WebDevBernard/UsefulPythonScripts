<h1 align="center">UsefulPythonScripts</h1>

This project contains various scripts I use at work.  I use it to automate the following:
- Create renewal letters using policy declarations
- Copy and paste ICBC policy documents
- Sort renewal list in Excel
- Delete broker copies in Intact home policy declarations

## Setup

```bash
pip install -r requirements.txt
python -m auto_py_to_exe
pip freeze | xargs pip uninstall -y`
```

## Pycharm
Copy idea.properties to `\Pycharm\bin`

```bash
{
    "terminal.integrated.profiles.windows": {
        "GitBash": {
            "source": "GitBash",
            "path": ["E:\\PortableGit\\bin\\bash.exe"],
            "icon": "terminal-bash"
          }
    },
    "terminal.integrated.defaultProfile.windows": "GitBash"
}

```