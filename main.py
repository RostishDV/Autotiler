from Executor import Executor
import json

def main():
    # set 1 if first time
    method = 1
    executor = Executor()
    if method == 1:
        executor.execute()
    if method == 2:
        executor.select_other_rows()


if __name__ == '__main__':
	main()