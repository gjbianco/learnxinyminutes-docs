---
language: excel
contributors:
    - ["Guy Bianco IV", "https://guy.sh"]
filename: spreadsheet.xml
---

Microsoft Excel has been a major component of the Microsoft Office suite of
programs for decades. Online and (mostly) offline versions are available on
nearly all major platforms and operating systems.

Although the application itself is not a programming language, the Excel
Formula Language is Turing-Complete and the application is by far the most
common way users interact with the language.

As such, Excel usage can range from storing raw data to creating entire
programs. However, the visual programming and interactivity of a spreadsheet
makes even a little understanding and knowledge go a long way. For example,
just knowing how a sheet works and the `SUM` function is enough to make a
robust, auto-updating personal-finance budget.

```
Files contain one or more spreadsheets.  Each spreadsheet is a vast
two-dimensional grid of cells. Each cell contains data or a formula.

Data can come in the form of strings and numbers. Data cells can be formatted,
including the display of the data itself.  For example, a numeric cell might
display as a currency with decimal places.

///////////////////////////////////

Formula cells begin with an equals sign (=) and automatically update as
cells referenced within the formula change. Formula cells can reference other
formula cells. In this way, formula cells can chain off of one another and
be used to break down otherwise-complex calculations.

A formula can reference a cell by its column letter and row number. For
example, the cell in column C and row 12 is referenced as C12.

Ranges are expressed as two cell references separated by a colon (:). For
example, the range from C12 to C26 would be C12:C26.

Formulas can include the regular mathematical operations addition (+),
subtraction (-), multiplication (*), and division (/). These can operate on
raw numbers in the formula as well as cell references. For example, the
following would take the value in (or result of) cell D2 and divide it by 7:

=D2/7
```

TODO

## Further Reading

The [Mozilla Developer Network][1] provides excellent documentation for
JavaScript as it's used in browsers. Plus, it's a wiki, so as you learn more you
can help others out by sharing your own knowledge.

MDN's [A re-introduction to JavaScript][2] covers much of the concepts covered
here in more detail. This guide has quite deliberately only covered the
JavaScript language itself; if you want to learn more about how to use
JavaScript in web pages, start by learning about the [Document Object Model][3].

[Learn Javascript by Example and with Challenges][4] is a variant of this
reference with built-in challenges.

[JavaScript Garden][5] is an in-depth guide of all the counter-intuitive parts
of the language.

[JavaScript: The Definitive Guide][6] is a classic guide and reference book.

[Eloquent Javascript][8] by Marijn Haverbeke is an excellent JS book/ebook with
attached terminal

[Javascript: The Right Way][10] is a guide intended to introduce new developers
to JavaScript and help experienced developers learn more about its best practices.

[javascript.info][11] is a modern javascript tutorial covering the basics (core language and working with a browser)
as well as advanced topics with concise explanations.

In addition to direct contributors to this article, some content is adapted from
Louie Dinh's Python tutorial on this site, and the [JS Tutorial][7] on the
Mozilla Developer Network.

[1]: https://developer.mozilla.org/en-US/docs/Web/JavaScript
[2]: https://developer.mozilla.org/en-US/docs/Web/JavaScript/A_re-introduction_to_JavaScript
[3]: https://developer.mozilla.org/en-US/docs/Using_the_W3C_DOM_Level_1_Core
[4]: http://www.learneroo.com/modules/64/nodes/350
[5]: http://bonsaiden.github.io/JavaScript-Garden/
[6]: http://www.amazon.com/gp/product/0596805527/
[7]: https://developer.mozilla.org/en-US/docs/Web/JavaScript/A_re-introduction_to_JavaScript
[8]: https://www.javascripttutorial.net/
[8]: http://eloquentjavascript.net/
[10]: http://jstherightway.org/
[11]: https://javascript.info/
