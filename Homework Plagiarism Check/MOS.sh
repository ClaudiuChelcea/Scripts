#!/bin/bash
# Author: Chelcea Claudiu-Marian

# Files extension
FILES_EXTENSION=$1

# File names
FIRST_FILE=$2
SECOND_FILE=$3

# Check for help command & arguments number
ARGUMENT_NUMBER=$#
if [[ $ARGUMENT_NUMBER -gt 3 ]]
then
	echo -e "Too many arguments!\nFormat is: ./mos_sh <extension> <file_1> <file_2>."
	exit -1
elif [[ $ARGUMENT_NUMBER -le 1 || $FILES_EXTENSION == "--help" ]]
then
	echo "Format: ./mos_sh <extension> <file_1> <file_2>."
	exit -1
elif [[ $ARGUMENT_NUMBER == 2 ]]
then
        SECOND_FILE=$FIRST_FILE
	FIRST_FILE=$FILES_EXTENSION
fi

# Check extension
if [[ $FIRST_FILE == $FILES_EXTENSION ]]
then
	# Check if files exists first
	if [[ -f $FIRST_FILE && -f $SECOND_FILE ]]
	then
		echo -e "No specified file extension!\nThe default file extension is '.c'\n"
		FILES_EXTENSION=c
	fi
fi

# Check files exist
if [[ $FIRST_FILE == "" || $SECOND_FILE == "" ||  ! -f $FIRST_FILE || ! -f $SECOND_FILE ]]
then
	# Throw error because we have no files
	echo -e "=============== WARNING ==============="
	echo -e "No files to check! Please specify two valid files!\nFormat: ./mos_sh <extension> <file_1> <file_2>."
	exit -1
else
	# Check extension
	EXTENSION_1=$(echo $1 | cut -d"." -f2)
	EXTENSION_2=$(echo $2 | cut -d"." -f2)

	if [[ $EXTENSION_1 != $FILES_EXTENSION || $EXTENSION_2 != $FILES_EXTENSION || $EXTENSION_1 !=  $EXTENSION_2 ]]
	then
		echo "Wrong files! Not all extensions match with" $FILES_EXTENSION!
		exit -1
	else
		# If the extension is correct, execute MOS
		perl ~/MOS/moss.pl -l $FILES_EXTENSION $FIRST_FILE $SECOND_FILE > tmp_file
		GET_LINK=$(cat tmp_file | tail -n 1)
		rm tmp_file
		
		# Get answer from link
		wget -O tmp_file $GET_LINK -q > /dev/null
		GET_ANSWER=$(cat tmp_file | tail -n 2| head -n 1)
		GET_FULL_ANSWER=$(cat tmp_file)
		rm tmp_file
		
		# Show answer
		echo $GET_ANSWER | grep "No matches were found in your submission." > /dev/null
		if [[ $? == 0 ]]
		then
			echo "No matches were found in your submission."
		else
			echo "Matches found:" $GET_LINK
			echo "Similarity percentage:" $(echo $GET_FULL_ANSWER | grep % | cut -d"%" -f1 | tail -n 1 | cut -d"(" -f2)%
		fi
	fi
fi
