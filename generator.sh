#!/bin/bash

set -x

day_of_today=$(date +%Y%m%d)
installation_dir="/home/sergio/azure-pricer/"
excel_files_dir=$installation_dir"output/"
max_number_of_xls_files=30

excel_file_of_today=$excel_files_dir"Azure-Quote-Tool-$day_of_today.xlsx"
readme_file_template=$excel_files_dir"README.MD.template"
readme_file=$installation_dir"README.MD"

cd $installation_dir

git pull
echo "UPDATING CODE FROM REPO"


python3 $installation_dir"xls_generator.py" $excel_file_of_today 

git config --global user.name "seryiogonzalez"

find $excel_file_of_today

if [ $? -ne 0 ]
then
	echo "ERROR"
	exit
fi

sed "s/__DATE__/$day_of_today/g" $readme_file_template > $readme_file

for old_file in $(find $excel_files_dir -type f -name "Azure-Quote-Tool-*.xlsx" -mtime +$max_number_of_xls_files)
do
	git rm $old_file
done

cd $installation_dir
git add $excel_file_of_today $readme_file
git commit -m "Automatic build of $day_of_today"
git push
