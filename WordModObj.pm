package WordModObj;

use strict;
use warnings;
use Data::Dumper;
use utf8;
binmode(STDOUT, ":utf8");
use open qw/:std :utf8/;

use Win32::OLE;
use Win32::OLE::Const 'Microsoft.Word';
use Win32::OLE::Const 'Microsoft Office';

use Win32::Console;

my $CONSOLE = Win32::Console->new(STD_OUTPUT_HANDLE);

$Win32::OLE::CP = Win32::OLE::CP_UTF8;
binmode STDOUT, 'encoding(CP932)';

sub new {
	my $class = shift;
	my %pars = (
		filename => undef,
		debug => undef,
		sleep => 0,
		@_
	);

	my $word = CreateObject Win32::OLE 'Word.Application' or die $!;
	$word->{'Visible'} = 1;

	my $document = $word->Documents->Add;
	$document->switch_view(view => "wdSeekMainDocument");
	my $selection = $word->Selection;

	my $search = $document->Content->Find;
	my $replace = $search->Replacement;

	my $hash = {
		sleep => $pars{sleep},
		filename => $pars{filename},
		debug => $pars{debug},
		word => $word,
		document => $document,
		selection => $selection,
		search => $search,
		replace => $replace
	};

	return bless $hash, $class;
}

sub get_filename {
	my $self = shift;
	$self->_debug("get_filename()");

	return $self->{filename};
}

sub insert_text {
	my $self = shift;
	my %pars = (
		text => '',
		noparagraph => 0,
		align => '',	### TODO!!!
		color => '',	### TODO!!!
		@_
	);

	$self->_debug("insert_text(text => $pars{text}, noparagraph => $pars{noparagraph}, align => $pars{align}, color => $pars{color})");

	my $aligned = 0;

	if($pars{align}) {
		my @valid_alignments = qw/
			wdAlignParagraphCenter
			wdAlignParagraphDistribute
			wdAlignParagraphJustify
			wdAlignParagraphJustifyHi
			wdAlignParagraphJustifyLow
			wdAlignParagraphJustifyMed
			wdAlignParagraphLeft
			wdAlignParagraphRight
			wdAlignParagraphThaiJustify
		/;

		if(grep($_ eq $pars{align}, @valid_alignments)) {
			$self->{selection}->ParagraphFormat->{Alignment} = eval "$pars{align}";
			$aligned = 1;
		} else {
			die "Invalid align option $pars{align}. Valid options:\n".join("\n", @valid_alignments)."\n";
		}
	}

	$self->{selection}->TypeText($pars{text});

	if($aligned) {
		$self->{selection}->ParagraphFormat->{Alignment} = wdAlignParagraphLeft;
	}

	if(!$pars{noparagraph}) {
		$self->{selection}->TypeParagraph;
	}

}

sub insert_heading {
	my $self = shift;
	my %pars = (
		level => 1,
		text => '',
		@_
	);

	$self->_debug("insert_heading(level => $pars{level}, text => $pars{text})");

	$self->{selection}->TypeText($pars{text});
	$self->{selection}->{'Style'} = "Heading $pars{level}";
	$self->{selection}->TypeParagraph;
}

sub edit_paragraph {
	my $self = shift;
	my %pars = (
		text => undef,
		paragraph => undef,
		@_
	);

	$self->_debug("edit_paragraph(text => $pars{text}, paragraph => $pars{paragraph})");
	$self->{document}->Paragraphs($pars{paragraph})->Range->{Text} = "$pars{text}\n";
}

sub get_paragraphs {
	my $self = shift;
	$self->_debug("get_paragraphs()");

	my @paragraphs = ();
	my $inputparagraphs = $self->{document}->Paragraphs;
	my $nparagraphs = $self->{document}->Paragraphs->Count;
	for my $i (1 .. $nparagraphs) {
		push @paragraphs, $inputparagraphs->Item($i)->Range->Text();
	}
	
	return map { s#\n\r#\n#g; $_; } @paragraphs;
}

sub replace_all {
	my $self = shift;
	my %pars = (
		oldtext => undef,
		newtext => undef,
		@_
	);
	$self->_debug("replace_all(oldtext => $pars{oldtext}, newtext => $pars{newtext})");
	$self->{search}->{Text} = $pars{oldtext};
	$self->{replace}->{Text} = $pars{newtext};
	$self->{search}->Execute({Replace => wdReplaceAll});
}

sub enter {
	my $self = shift;

	$self->_debug("enter()");
	$self->{document}->ActiveWindow->Selection->TypeParagraph;
}

sub switch_view {
	# use switch_view to change to header, footer, main document and so on...

	my $self = shift;
	my %pars = (
		view => undef,
		@_
	);
	my @possible_views = qw/
		wdSeekCurrentPageFooter
		wdSeekCurrentPageHeader 
		wdSeekEndnotes 
		wdSeekEvenPagesFooter 
		wdSeekEvenPagesHeader 
		wdSeekFirstPageFooter 
		wdSeekFirstPageHeader 
		wdSeekFootnotes 
		wdSeekMainDocument 
		wdSeekPrimaryFooter 
		wdSeekPrimaryHeader 
	/;

	$self->_debug("switch_view(view => $pars{view})");

	if(grep($_ eq $pars{view}, @possible_views)) {
		$self->{document}->ActiveWindow->ActivePane->View->{SeekView} = $pars{view};
	} else {
		die "Wrong view: $pars{view}. Possible views:\n".join("\n", @possible_views)."\n";
	}
}

sub save_doc_as {
	my $self = shift;
	my %pars = (
		filename => undef,
		@_
	);

	$self->_debug("save_doc_as(filename => $pars{filename})");

	$self->{document}->SaveAs($pars{filename});
}

sub close_doc {
	my $self = shift;
	$self->_debug("close_doc()");
	$self->{document}->Close;
}

sub insert_page_break { ### TODO
	my $self = shift;
	$self->_debug("insert_page_break()");
	$self->{document}->ActiveWindow->Selection->{Range}->InsertBreak(wdPageBreak);
}

sub _debug {
	my $self = shift;

	if($self->{debug}) {
		foreach (@_) {
			$\ = '';
			my $attr = $CONSOLE->Attr(); # Get current console colors
			$CONSOLE->Attr($FG_YELLOW | $BG_GREEN);
			print "DEBUG: $_\n";
			$CONSOLE->Attr($attr);
			$\ = "\n";
			if($self->{sleep}) {
				my $attr = $CONSOLE->Attr(); # Get current console colors
				$CONSOLE->Attr($FG_YELLOW | $BG_GREEN);
				print "Sleeping $self->{sleep} seconds...\n";
				$CONSOLE->Attr($attr);
				$\ = "\n";
				sleep $self->{sleep};
			}
		}
	}
}

1;
