example-ctf:
	poetry run python cost_calculator/cli.py -ctf example/cost_table_files example/fca_files_empty

example-ftb:
	poetry run python cost_calculator/cli.py -ftb example/fca_files example/BrakeSystem_BOM.xlsx

example-stf:
	poetry run python cost_calculator/cli.py -stf example/fca_files