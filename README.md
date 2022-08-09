# daoatlasExcel
用户daoatlas数据excel整理


# 开发

创建虚拟环境
python3 -m venv .env
删除所有pip包
#pip freeze --all | xargs pip uninstall -y

# 导出安装包

pip freeze > requirements.txt

# 安装

pip install -r requirements.txt

# 运行
python main.py
