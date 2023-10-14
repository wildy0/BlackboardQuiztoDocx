# BlackboardQuiztoDocx
This project can read Quiz/assessment zip files from Blackboard LMS exports.  The zip file is processed, and the quiz questions are output to a word docx file.  The word files will contain Questions (essay questions) and answers (if given in the quiz setup).  MCQs have correct answer highlighted and a separate output LAMs TBL word import format.

The zip file will be parsed, and output will be added to a directory of the same name as the input file.

Occasionally you may run in to long filename issues for output, if this happens shorten the filename of the downloaded zip.

The script does not handle errors cleanly.

The script creates a directory with unzipped content which it deletes after use.  If a directory already exists with that name, then it will be deleted; use with care.

The script outputs docx for each data file within the zip separately.  For question banks there is normally a separate datafile in the zip export for each bank of a quiz.  You will get an output file for each bank.  For quiz/assessments which contain banks the script will also output the top level quiz/assessment which will contain a table of all the banks used and the number of questions the quiz takes from each bank by bank name corresponding to the docx file output name in the output directory.
