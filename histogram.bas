Sub histogramVysky()
'
'creating histogram of dimensions of books using Excel function
'
    Range("AB16").Value = "5" 'creating values for histogram. Could be replaced by FOR
    Range("AB17").Value = "10"
    Range("AB18").Value = "15"
    Range("AB19").Value = "20"
    Range("AB20").Value = "25"
    Range("AB21").Value = "30"
    Range("AB22").Value = "35"
    Range("AB23").Value = "40"
    Application.Run "ATPVBAEN.XLAM!Histogram", ActiveSheet.Range("$V$3:$V$1000") _
        , ActiveSheet.Range("$AC$15"), ActiveSheet.Range("$AB$16:$AB$23"), False, _
        False, False, False
    Application.Run "ATPVBAEN.XLAM!Histogram", ActiveSheet.Range("$W$3:$W$1000") _
        , ActiveSheet.Range("$AC$26"), ActiveSheet.Range("$AB$16:$AB$23"), False, _
        False, False, False
        'calling histogram function of Excel

        
    Range("AC15:AC35").NumberFormat = "@"   'creating labels for histogram
    Range("AC16:AC24").HorizontalAlignment = xlRight
    Range("AC27:AC35").HorizontalAlignment = xlRight
    Range("AC16").Value = "0 - 5"
    Range("AC17").Value = "5 - 10"
    Range("AC18").Value = "10 - 15"
    Range("AC19").Value = "15 - 20"
    Range("AC20").Value = "20 - 25"
    Range("AC21").Value = "25 - 30"
    Range("AC22").Value = "30 - 35"
    Range("AC23").Value = "35 - 40"
    Range("AC24").Value = "<40"

    Range("AC27").Value = "0 - 5"
    Range("AC28").Value = "5 - 10"
    Range("AC29").Value = "10 - 15"
    Range("AC30").Value = "15 - 20"
    Range("AC31").Value = "20 - 25"
    Range("AC32").Value = "25 - 30"
    Range("AC33").Value = "30 - 35"
    Range("AC34").Value = "35 - 40"
    Range("AC35").Value = "<40"
    
    Range("AC15").Value = "Dimension"
    Range("AC26").Value = "Dimension"
    Range("AD15").Value = "Amount of b."
    Range("AD26").Value = "Amount of b."
End Sub
