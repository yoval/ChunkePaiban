conda create --name paiban python=3.7

conda activate paiban

pip install PySimpleGUI -i https://pypi.tuna.tsinghua.edu.cn/simple

pip install pandas -i https://pypi.tuna.tsinghua.edu.cn/simple

pip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple

pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple

pyinstaller -F -w paiban.py