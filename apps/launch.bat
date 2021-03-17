"%java8home%/bin/java" -jar excel-diff-checker.jar -b old.* -t new.* %* > edc.log
pause

"%java8home%/bin/java" -jar excel-diff-checker.jar -b new.* -t old.* %* > edc-opp.log
pause