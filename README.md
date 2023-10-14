# BlackboardQuiztoDocx
This project can read Quiz/assessment zip files from Blackboard LMS exports.  These are the output zip files which contain Quiz Questions and answers. This is handy to keep copies of your quizes/MCQs etc outside of the LMS and send to others including external examiners for example. (carefully check the content as output might go wrong, or be missed because of the manual html to docx conversion).

The zip file is processed, and the quiz questions are output to a word docx file.  The word files will contain Questions (essay questions) and answers (if given in the quiz setup).  MCQs have correct answer highlighted and a separate output LAMs TBL word import format.

The zip file will be parsed, and output will be added to a directory of the same name as the input file.

Occasionally you may run in to long filename issues for output, if this happens shorten the filename of the downloaded zip.

The script does not handle errors cleanly.

The script creates a directory with unzipped content which it deletes after use.  If a directory already exists with that name, then it will be deleted; use with care.

The script outputs docx for each data file within the zip separately.  For question banks there is normally a separate datafile in the zip export for each bank of a quiz.  You will get an output file for each bank.  For quiz/assessments which contain banks the script will also output the top level quiz/assessment which will contain a table of all the banks used and the number of questions the quiz takes from each bank by bank name corresponding to the docx file output name in the output directory.

The template.docx file is needed for the MCQ quiz output to ensure the lists a) b) c) d) etc format can be added to the output docx for MCQs.

The htmltodocx.py code converts the html in the blackboard quiz export to the word file.  Not all HTML elements will be perfectly preserved.  Font changes and other formatting will likely completely fail, and you may end up with missing output in the quiz.  Use with care.  Tables and some limited formatting should work.
For MCQ the correct answer will be coloured RED.
For MCQ and essay answers the correct answer in the quiz (for the markers) will appear after each question and will be coloured RED.

For MCQs there will be two output, one with answers in red and containing answer comments/feedback.  The other will be _LAMS for import in to the LAMS word format Quiz import for LAMS TBL.  You should be able to use this script to quickly get large banks/quiz in to LAMS.  It may not always work.  I had to hackaround in the xml of docx to get the lists for MCQs to restart numbering for each question.  The created output is a bodged version of a docx which does not properly follow docx file standards but it seems to work ok in word.  It is unclear if importing programmes will read it properly or not.
