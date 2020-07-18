# Number to Words in Excel or Google Sheets
 MS Excel or Google Sheets Formula that converts Numbers into Words. No VBA Coding is required.


```ruby
=TRIM(IF(OR(LEN(FLOOR(A10,1))=13,FLOOR(A10,1)<=0),"-nill-",PROPER(SUBSTITUTE(CONCATENATE(
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),1,1)+1,"","one hundred ","two hundred ","three hundred ","four hundred ","five hundred ","six hundred ","seven hundred ","eight hundred ","nine hundred "),
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),2,1)+1,"",
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),3,1)+1,"ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"),"twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"),IF(VALUE(MID(TEXT(INT(A10),REPT(0,12)),2,1))>1,
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),3,1)+1,"","-one","-two","-three","-four","-five","-six","-seven","-eight","-nine"),IF(VALUE(MID(TEXT(INT(A10),REPT(0,12)),2,1))=0,
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),3,1)+1,"","one","two","three","four","five","six","seven","eight","nine"),"")),IF(A10>=10^9," billion ",""),
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),4,1)+1,"","one hundred ","two hundred ","three hundred ","four hundred ","five hundred ","six hundred ","seven hundred ","eight hundred ","nine hundred "),
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),5,1)+1,"",
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),6,1)+1,"ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"),"twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"),IF(VALUE(MID(TEXT(INT(A10),REPT(0,12)),5,1))>1,
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),6,1)+1,"","-one","-two","-three","-four","-five","-six","-seven","-eight","-nine"),IF(VALUE(MID(TEXT(INT(A10),REPT(0,12)),5,1))=0,
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),6,1)+1,"","one","two","three","four","five","six","seven","eight","nine"),"")),IF(VALUE(MID(TEXT(INT(A10),REPT(0,12)),4,3))>0," million ",""),
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),7,1)+1,"","one hundred ","two hundred ","three hundred ","four hundred ","five hundred ","six hundred ","seven hundred ","eight hundred ","nine hundred "),
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),8,1)+1,"",
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),9,1)+1,"ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"),"twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"),IF(VALUE(MID(TEXT(INT(A10),REPT(0,12)),8,1))>1,
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),9,1)+1,"","-one","-two","-three","-four","-five","-six","-seven","-eight","-nine"),IF(VALUE(MID(TEXT(INT(A10),REPT(0,12)),8,1))=0,
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),9,1)+1,"","one","two","three","four","five","six","seven","eight","nine"),"")),IF(VALUE(MID(TEXT(INT(A10),REPT(0,12)),7,3))," thousand ",""),
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),10,1)+1,"","one hundred ","two hundred ","three hundred ","four hundred ","five hundred ","six hundred ","seven hundred ","eight hundred ","nine hundred "),
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),11,1)+1,"",
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),12,1)+1,"ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"),"twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"),IF(VALUE(MID(TEXT(INT(A10),REPT(0,12)),11,1))>1,
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),12,1)+1,"","-one","-two","-three","-four","-five","-six","-seven","-eight","-nine"),IF(VALUE(MID(TEXT(INT(A10),REPT(0,12)),11,1))=0,
CHOOSE(MID(TEXT(INT(A10),REPT(0,12)),12,1)+1,"","one","two","three","four","five","six","seven","eight","nine"),"")))," "," ")&IF(FLOOR(A10,1)>1," Rupees Only."," Peso"))&IF(ISERROR(FIND(".",A10,1))," "," and "&PROPER(IF(LEN(LEFT(TRIM(MID(SUBSTITUTE(A10,".",REPT(" ",255)),255,200)),2))=1,
CHOOSE(1*LEFT(TRIM(MID(SUBSTITUTE(A10,".",REPT(" ",255)),255,200)),2),"ten","twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety")&" ","")&CONCATENATE(
CHOOSE(MID(TEXT(INT(LEFT(TRIM(MID(SUBSTITUTE(A10,".",REPT(" ",255)),255,200)),2)),REPT(0,12)),11,1)+1,"",
CHOOSE(MID(TEXT(INT(LEFT(TRIM(MID(SUBSTITUTE(A10,".",REPT(" ",255)),255,200)),2)),REPT(0,12)),12,1)+1,"ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen")&" ","twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"),IF(VALUE(MID(TEXT(INT(LEFT(TRIM(MID(SUBSTITUTE(A10,".",REPT(" ",255)),255,200)),2)),REPT(0,12)),11,1))>1,
CHOOSE(MID(TEXT(INT(LEFT(TRIM(MID(SUBSTITUTE(A10,".",REPT(" ",255)),255,200)),2)),REPT(0,12)),12,1)+1,"","-one","-two","-three","-four","-five","-six","-seven","-eight","-nine")&" ",IF(LEFT(TRIM(MID(SUBSTITUTE(A10,".",REPT(" ",255)),255,200)),2)="01","one cent",IF(LEFT(TRIM(MID(SUBSTITUTE(A10,".",REPT(" ",255)),255,200)),1)="0",
CHOOSE(MID(TEXT(INT(LEFT(TRIM(MID(SUBSTITUTE(A10,".",REPT(" ",255)),255,200)),2)),REPT(0,12)),12,1)+1,"","one","two","three","four","five","six","seven","eight","nine")&" ",""))))))))
```
