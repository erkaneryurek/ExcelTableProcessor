# ExcelTableProcessor
This program turns the output of a course scheduling program into database format.
The excel file chosen must have headers.
The format of the excel file should be like below;

- COURSE_CODE	DAY	SLOT
- BİLP1111.1  MMM 123
- BİLP1111.2  TTT 123
- TURK1201.1  MM  45
- .
- .
- .

The output will be like;

- BİLP1111.1  M1
- BİLP1111.1  M2
- BİLP1111.1  M3
- BİLP1111.2  T1
- BİLP1111.2  T2
- BİLP1111.2  T3
- TURK1201.1  M4
- TURK1201.1  M5
- .
- .
- .

For day formats M, T, W, F are acceptable but Th is not because it is 2 characters. What you can do is select DAY column on your excel and replace all Th's with P 
and after you run the program replace all the P's with Th's.

For slots higher than 9 like with Th you can select SLOT column on your excel and replace all 10, 11, 12 with A, B, C repectively and 13, 14, 15 with X, Y, Z
and after you run the program replace all the A, B, C with 10, 11, 12 and X, Y, Z with 13, 14, 15.

That is all.
