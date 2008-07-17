#!/usr/bin/perl

#
# luciferase-collator
#   collates intensity spreadsheet(s) with a library spreadsheet to bring
#   together all the information for the luciferase experiments with the
#   RNAi kit.
# author:  Julian Selley <j.selley@manchester.ac.uk>
# created: 29-May-2008
# example: src/luciferase-collator.pl -i data/intensity-Elk-Vp16.xls \
#            -l data/library-haslam\(sharrocks\)/mouse\ protein\ kinases.xls \
#            -o data/collation-haslam\(sharrocks\)/mouse\ protein\ kinases.xls
#

# Example of handling Excel spreadsheets pieced together from examples on:
#   http://www.ibm.com/developerworks/linux/library/l-pexcel/

#use strict;
#use warnings;

use Data::Dumper;
use Getopt::Long;
use Spreadsheet::ParseExcel;
use Spreadsheet::WriteExcel;

## VARIABLES ###################################################################
our %opt;  # command-line options
our %inten;  # hash of intensity matricies
our $ref_library;  # library matrix
our $ref_collation;  # collation of desired info to write to spreadsheet output


## MAIN ########################################################################
# get options
GetOptions('help|h'         => \$opt{'help'},          # help/usage
           'intensity|i=s@' => \$opt{'fn_intensity'},  # intensity spreadsheet filename
           'library|l=s'    => \$opt{'fn_library'},    # library spreadsheet filename
           'output|o=s'     => \$opt{'fn_output'});    # output spreadsheet filename
usage() if $opt{'help'};

# exit if spreadsheets not specified or don't exist
die "ERROR: unable to load the intensity spreadsheet\n"
  if (not defined $opt{'fn_intensity'} || ! -e $opt{'fn_intensity'});
die "ERROR: unable to load the library spreadsheet\n"
  if (not defined $opt{'fn_library'} || ! -e $opt{'fn_library'});
die "ERROR: output spreadsheet not specified\n"
  if (not defined $opt{'fn_output'});


# load the intensity spreadsheet(s)
for (my $int_fni = 0; $int_fni < @{$opt{'fn_intensity'}}; $int_fni++) {
  my ($ref_intensity, $date, $time, $instr_bg) =
    loadIntensities(${$opt{'fn_intensity'}}[$int_fni], 14, 1);
  # generate key for indexing intensity
  my $_key = "$date,$time";
  $_key .= (defined $instr_bg) ? ",$instr_bg" : "";

  # store the intensity matrix in the hash of intensity matricies
  $inten{$_key} = $ref_intensity;
}
# load the library spreadsheet
$ref_library = loadLibrary($opt{'fn_library'});

# collate the library and intensities
$ref_collation = collateInfo($ref_library, \%inten);

# write the collated information out
our $_dest_book = new Spreadsheet::WriteExcel($opt{'fn_output'})
  or die "ERROR: unable to create a new Excel spreadsheet: $!";
our $_dest_sheet = $_dest_book->addworksheet("Sheet 1");

for (my $colr = 0; $colr < @{$ref_collation}; $colr++) {
  my @_collr = @{${$ref_collation}[$colr]};
  for (my $colc = 0; $colc < @_collr; $colc++) {
    $_dest_sheet->write($colr, $colc, $_collr[$colc]);
  }
}

$_dest_book->close();



## FUNCTIONS ###################################################################
# collateInfo
#   collates the library info and the intensities
# arg: library matrix - the array of data produced from loading the library
#        spreadsheet
#      hash of intensity array's - the arrays produced by loading the intensity
#        spreadsheet(s) -keyed on the date, time and instr_bg
# ret: a matrix: the spreadsheet of the collated information
sub collateInfo {
  # argument retreival
  my $ref_lib = shift;
  my $ref_int_matrix = shift;

  # variables
  my @library = @{$ref_lib};
  my %intensities_matrix = %{$ref_int_matrix};
  my @collation;

  # put in the titles of the columns
  push @collation, ["Barcode", "Plate #", "Well", "Gene Name", "Gene ID", 
                    "Accession", "GI Number", keys %intensities_matrix];

  for (my $libi = 0; $libi < @library; $libi++) {
    my %_lib_row = %{$library[$libi]};

    # create a temporary array to put the row data in
    #  - start with library information
    my @_tmp = ($_lib_row{'barcode'},
                $_lib_row{'plate'},
                $_lib_row{'well'},
                $_lib_row{'name'},
                $_lib_row{'id'},
                $_lib_row{'acc'},
                $_lib_row{'gi'});

    #  - add the intensity matrix info
    foreach my $ref_int (keys %intensities_matrix) {
      my @int = @{$intensities_matrix{$ref_int}};

      push @_tmp, $int[$_lib_row{'wellr'}][$_lib_row{'wellc'}];
    }

    # push the row to collation
    push @collation, \@_tmp;
  }

  # return the collated results
  return \@collation;
}

# loadIntensities
#   loads the intensity spreadsheet specified
# arg: spreadsheet_fn - spreadsheet filename
#      start_row - where the intensities start (assuming a matrix layout)
#      start_col - where the intensities start (assuming a matrix layout)
# ret: ref. to an array of intensity values
#      the date & time of the intensity background
#      the "instrument background" naming
sub loadIntensities {
  # argument retreival
  my $spreadsheet_fn = shift;  # spreadsheet filename
  my $start_row = shift;  # row where the intensities start (assuming a matrix layout)
  my $start_col = shift;  # col where the intensities start (assuming a matrix layout)

  # variables
  my $src = new Spreadsheet::ParseExcel;
  my $src_book = $src->Parse($spreadsheet_fn);
  my $src_sheet = $src_book->{Worksheet}[0];  # take the first worksheet from the book
  my ($date, $time);  # the expr. date and time
  my $instr_bg;  # the "instrument background" information
  my @intensities;  # the intensities

  # get the expr date and time
  $date = $src_sheet->{Cells}[2][1]->{'_Value'};
  $time = $src_sheet->{Cells}[3][1]->{'_Value'}; $time =~ s/ AM\/PM//;
  # get the instrument background details
#  $instr_bg = $src_sheet->{Cells}[6][0]->Value;

  # get the intensities
  for (my $irow = $start_row; $irow <= $src_sheet->{MaxRow}; $irow++) {
    for (my $icol = $start_col; $icol <= $src_sheet->{MaxCol}; $icol++) {
      $intensities[$irow - $start_row][$icol - $start_col] =
        $src_sheet->{Cells}[$irow][$icol]->Value;
    }
  }

  # return intensities, date, time, instr_bg
  return (\@intensities, $date, $time, $instr_bg);
}

# loadLibrary
#   loads the library spreadsheet specified
# arg: spreadsheet_fn - spreadsheet filename
# ret: ref. to the library information
sub loadLibrary {
  # argument retreival
  my $spreadsheet_fn = shift;  # spreadsheet filename

  # variables
  my $src = new Spreadsheet::ParseExcel;
  my $src_book = $src->Parse($spreadsheet_fn);
  my $src_sheet = $src_book->{Worksheet}[0];  # take the first worksheet from the book
  my @library;
  my $_last_well = "";

  for (my $row = $src_sheet->{MinRow}; $row <= $src_sheet->{MaxRow}; $row++) {
    next unless (defined $src_sheet->{Cells}[$row][2] &&  # move on if a blank row
      $src_sheet->{Cells}[$row][2]->Value =~ /^Plate \d+/);  # or the row doesn't begin with 'Plate n'
    next if ($src_sheet->{Cells}[$row][3]->Value eq $_last_well);

    my %rowinfo = ();

    # get plate information
    $rowinfo{'plate'} = $src_sheet->{Cells}[$row][2]->Value;  # plate number
    $rowinfo{'plate'} =~ s/^Plate //;  # tidy up the plate number

    # get row information
    $rowinfo{'well'}  = $src_sheet->{Cells}[$row][3]->Value;  # get the well location
    $_last_well = $rowinfo{'well'};
    # identify the components of the well location
    my $_well = $src_sheet->{Cells}[$row][3]->Value;
    $_well =~ /([A-Z])(\d+)/;
    $rowinfo{'wellr'} = ord($1) - 65;
    $rowinfo{'wellc'} = $2 - 1;

    # get the barcode
    $rowinfo{'barcode'} = $src_sheet->{Cells}[$row][1]->Value;  # the barcode
    # get gene information
    $rowinfo{'name'} = $src_sheet->{Cells}[$row][6]->Value  # the gene name
      if($src_sheet->{Cells}[$row][6]);
    $rowinfo{'id'}   = $src_sheet->{Cells}[$row][7]->Value  # the gene id
      if($src_sheet->{Cells}[$row][7]);
    $rowinfo{'acc'}  = $src_sheet->{Cells}[$row][8]->Value  # the gene accession
      if($src_sheet->{Cells}[$row][8]);
    $rowinfo{'gi'}   = $src_sheet->{Cells}[$row][9]->Value  # the gi number
      if($src_sheet->{Cells}[$row][9]);

    # store the row information
    push @library, \%rowinfo;
  }

  # return the library information
  return \@library;
}

# usage
#   prints the usage information for the program (a short help)
# arg: none
# ret: none
sub usage {
  print STDERR "$0 [-h] -i <intensity_spreadsheet> -l <library_spreadsheet> -o <output_spreadsheet>\n";
  print STDERR "  -h|--help:      this information\n";
  print STDERR "  -i|--intensity: the Excel spreadsheet containing the intensities\n";
  print STDERR "  -l|--library:   the Excel spreadsheet containing the library info\n";
  print STDERR "  -o|--output:    the Excel spreadsheet generated\n";
}
