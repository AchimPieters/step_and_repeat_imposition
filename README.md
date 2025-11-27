Step-and-Repeat Imposition
==========================

Python tool to impose business cards or small print layouts multiple times on a single sheet
(A4, A3, SRA4, SRA3), including duplex correction for front/back misalignment.


FEATURES
--------

- Automatically calculates the maximum number of cards per sheet
- Supports A4, A3, SRA4 and SRA3 paper sizes
- Automatically tries both 0 mm and 2 mm trimming
- Optional rotation for better fit
- Duplex printing correction for backside misalignment
- Grid is always centered inside the printable area
- Supports asymmetric and symmetric printer margins


REQUIREMENTS
------------

Python 3.8 or higher

Install dependency:

pip install pypdf

(Falls back to PyPDF2 if pypdf is not installed)


USAGE
-----

Basic:

python step_and_repeat_imposition.py input.pdf

Specify paper size:

python step_and_repeat_imposition.py input.pdf --paper A3

Symmetric margin on all sides:

python step_and_repeat_imposition.py input.pdf --margin-mm 5

Separate margins:

python step_and_repeat_imposition.py input.pdf --margin-x-mm 5 --margin-y-mm 8

Specify output file:

python step_and_repeat_imposition.py input.pdf output.pdf


INPUT FORMAT
------------

The input PDF must contain at least two pages:

Page 1 = front  
Page 2 = back


DUPLEX ALIGNMENT
----------------

Backside correction is set in the script:

BACK_OFFSET_X_MM = -2.5
BACK_OFFSET_Y_MM =  0.0

Positive Y moves the back side upward relative to the front.
Positive X moves the back side to the right relative to the front.


SUPPORTED PAPER SIZES
---------------------

A4    = 210 x 297 mm  
A3    = 297 x 420 mm  
SRA4  = 225 x 320 mm  
SRA3  = 320 x 450 mm  


EXAMPLE
-------

Example for SRA3 with 5 mm margins:

python step_and_repeat_imposition.py cards.pdf --paper SRA3 --margin-mm 5

Output file:

cards_PRINT.pdf


LICENSE
-------

MIT License


AUTHOR
------

Achim Pieters
