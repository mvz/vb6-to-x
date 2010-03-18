VB6-to-X
========

Convert VB6 projects to language/platform X, for some values of X.

VB6-to-X will read vbp files and convert the whole project, including
forms, modules and resources.

Ruby will be the first supported value of X. Python may be a good candidate
for the second.

Status
------

It's early days. The parser can parse one particular .frm file. 

Design
------

VB6-to-X will consist of

1. A parser generating an AST
2. A generic module for walking the AST.
3. A X-specific module implementing callbacks for emitting target language
   code.
4. Miscellaneous bits for stuff like creating the target files and extracting
   resources from the .frx files.

VB6-to-X uses Treetop as its parser framework.

License
-------

VB6-to-X will be licensed under the GPL.
