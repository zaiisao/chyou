#!/usr/bin/python
# -*- coding: iso-8859-15 -*-

import argparse
import sys
import pandas as pd
import mariadb

def main():
	parser = argparse.ArgumentParser()
	parser.add_argument('file', help='excel file')

	args = parser.parse_args()
	file = args.file

	df = pd.read_excel(file)
	df = df.iloc[7:]

	people = []

	for row in df.iterrows():
		name = row[1][2]
		maeul = row[1][14] if row[1][14] != 0 else None
		mokjang = row[1][15] if row[1][14] != 0 else None

		people.append({
			'name': name,
			'maeul': maeul,
			'mokjang': mokjang
		})

	password = input("pw: ")

	# Connect to MariaDB Platform
	try:
	    conn = mariadb.connect(
	        user="juyoung32",
	        password=password,
	        host="chyou.co.kr",
	        port=3306#,
	        #database="employees"

	    )
	except mariadb.Error as e:
	    print(f"Error connecting to MariaDB Platform: {e}")
	    sys.exit(1)

if __name__ == '__main__':
    main()