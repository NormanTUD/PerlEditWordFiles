# PerlEditWordFiles

This allows editing word files via a Perl-Script on Windows (needs english Word 2016+ to work) 

## Dependencies

This needs Strawberry Perl, Win32::OLE, Win32::OLE::Const, Win32::Console and Data::Dumper to be installed
You can install each of these (except for Strawberry Perl) with

> cpan -i Name::Of::Module

Also, you need Office installed.

## Run it

You can run the test-script by simply running

> perl autoobj.pl

## Features

- Adding Lines
- Adding Headings
- Changing paragraphs (e.g. when a regex matches, see autoobj.pl for an example)
- Adding new paragraphs
- "Replace all" feature
- Get the whole text (as an array of paragraphs)

## Example video

Here you can see an example video (with sleeping enabled between commands so it's more easily visible what it does):

https://youtu.be/02D_K5qyb50

