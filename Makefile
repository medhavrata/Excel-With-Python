virtual:
	python3 -m venv ~/.Excel-With-Python

source:
	source ~/.Excel-With-Python/bin/activate

run:
	python openexcel.py

install:
	# pip install --upgrade pip &&\
	pip install -r requirements.txt

test-initial:
	python -m pytest -vv --cov=hello tests
	#python -m pytest --nbval notebook_test.ipynb	#tests our jupyter notebook

test:
	python -m pytest -vvv --cov=hello --cov=greeting \
		--cov=smath --cov=web tests
	python -m pytest --nbval notebook.ipynb	#tests our jupyter notebook
	#python -m pytest -v tests/test_web.py #if you just want to test web

debug:
	python -m pytest -vv --pdb	#Debugger is invoked

one-test:
	python -m pytest -vv tests/test_greeting.py::test_my_name1

debugthree:
	#not working the way I expect
	python -m pytest -vv --pdb --maxfail=3  # drop to PDB for first three failures

format:
	black *.py

lint:
	pylint --disable=R,C hello.py

all: install lint test-initial format