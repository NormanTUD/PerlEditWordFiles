use warnings;
use strict;
use Data::Dumper;
use lib '.';
$\ = "\n";
use WordModObj;
use utf8;

my $document = WordModObj->new(filename => "E:\\test.docx", debug => 1, sleep => 0);

$document->goto_top();
$document->enter();
$document->goto_top();
$document->insert_heading(text => "Heading 1", level => 1);
$document->insert_heading(text => "Heading 2", level => 2);
$document->insert_text(text => "Replace me!");
$document->insert_text(text => "Remove Me. Keep me!");
$document->enter();
$document->insert_text(text => "Test 123");
$document->insert_text(text => "Test 456");
$document->insert_text(text => "Test ".rand().", ".rand());
#$document->insert_page_break();
$document->insert_text(text => "test new page");
$document->replace_all(oldtext => "Heading", newtext => "Überschrift");
$document->edit_paragraph(paragraph => 3, text => "Replaced Paragraph 123");
$document->add_comment(paragraph => 3, text => "Testkommentar");


my @paragraphs = $document->get_paragraphs();
foreach my $paragraph_index (0 .. $#paragraphs) {
	my $text = $paragraphs[$paragraph_index];
	if($text) {
		if($text =~ m#Remove Me\. #) {
			$text =~ s#Remove Me\. ##g;
			$document->edit_paragraph(paragraph => $paragraph_index + 1, text => $text);
		}
		
		if($text =~ m#Test#) {
			$document->add_comment(paragraph => $paragraph_index + 1, text => "Testkommentar 2");
		}
	}
}

print ">>>>>\n".join("\n", $document->get_paragraphs())."\n<<<<<\n";

$document->save_doc_as();
#$document->close();
#$document->quit();
