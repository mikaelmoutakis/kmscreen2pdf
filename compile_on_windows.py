from subprocess import run

cmd1 = ["pip", "install", "-r", "requirements.txt"]
run(cmd1, check=True)
cmd2 = ["pyinstaller", "--onefile", "kmscreen2pdf/kmscreen2pdf.py"]
run(cmd2, check=True)
