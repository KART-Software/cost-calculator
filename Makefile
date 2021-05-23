example-ctf:
	poetry run python cost_calculator/cli.py -ctf example/cost_table_files example/fca_files_empty

example-ftb:
	poetry run python cost_calculator/cli.py -ftb example/fca_files example/BrakeSystem.xlsx