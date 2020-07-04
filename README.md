# excel-diff-checker
![](https://img.shields.io/badge/Release-V1.0.0-blue.svg) ![](https://img.shields.io/badge/Build-Stable-green.svg) [![License](https://img.shields.io/badge/License-Apache%202.0-red.svg)](https://opensource.org/licenses/Apache-2.0) ![](https://img.shields.io/badge/By-Abhishek%20Sarkar-red.svg?style=social&logo=appveyor)

------------


This is a simple java tool to check diff between two excel files.

## Prerequisites
Java 8+

## Usage
java -jar excel-diff-checker.jar \<File1-path> \<File2-path> [r] -> where File1 and File2 are mandatory, option: 'r' is optional.

## Notes to follow
 - This tool is not having any complex algorithm to check diff, so won't be able to detect column/row addition/deletion.
 - It basically considers File1 as base, and checks for diff cell-by-cell in File2, even for sheets too.
 - By default, it adds cell comments/note saying like 'Expected: value1, Found: value2' in a copy of File1 excel and produces it as a different result file.
 - Instead of above, if only remarks required about different cell diffs i.e. no separate reult file required, in that case, the option: 'r' can be used, which just prints the diff note as mentioned in above point, in the console output and not as cell comment of a new file.
