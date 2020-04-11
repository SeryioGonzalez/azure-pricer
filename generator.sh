#!/bin/bash

set -x

dayOfToday=$(date +%Y%m%d)
installationDir="/home/sergio/azure-pricer/"
excelFilesDir=$installationDir"output/"
daysToDelete=30

excelFileOfToday=$excelFilesDir"Azure-Quote-Tool-$dayOfToday.xlsx"
readMeFileTemplate=$excelFilesDir"README.MD.template"
readMeFile=$installationDir"README.MD"

cd $installationDir

git pull
echo "UPDATING CODE FROM REPO"


python3 $installationDir"xls_generator.py" $excelFileOfToday 

git config --global user.name "seryiogonzalez"

find $excelFileOfToday

if [ $? -ne 0 ]
then
	echo "ERROR"
	exit
fi

sed "s/__DATE__/$dayOfToday/g" $readMeFileTemplate > $readMeFile

for oldFile in $(find $excelFilesDir -type f -name "Azure-Quote-Tool-*.xlsx" -mtime +$daysToDelete)
do
	git rm $oldFile
done

cd $installationDir
git add $excelFileOfToday $readMeFile
git commit -m "Automatic build of $dayOfToday"
git push
