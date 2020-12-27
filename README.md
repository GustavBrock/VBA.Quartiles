# VBA.Quartiles
![Help](https://raw.githubusercontent.com/GustavBrock/VBA.Quartiles/master/images/EE%20Header.png)

### Calculate quartiles in 20 ways and medians in Microsoft Access
There is no single "correct" calculation method for quartiles. Excel features two methods, some math and statistic packages offer some more, but here is presented no less than twenty methods for various purposes using VBA and Microsoft Access.

To calculate a quartile of a sample is in theory easy, and is much like calculating the median. The difficult part is the implementation; contrary to calculating the median, there exists no single specific method that stands above the rest or can be considered the "best" method among the about twenty known methods for calculating a quartile. The "best" method will be the method that fits the purpose or - in some areas - is considered a de-facto standard.

### Methods
It is quite hard to even obtain a list of known methods for calculating a quartile, not to say proven results from these. 
The methods have been collected as an enum including as in-line comments their names, applications, and sources, together with their basic calculation methods for the first and the third quartile (the second is always calculated as the median):

```
' Quartile calculation methods.
' Values equal those listed in the source. See function Quartile.
'
' Common names of variables used in calculation formulas.
'
' L: Q1, Lower quartile.
' H: Q3, Higher quartile.
' M: Q2, Median (not used here).
' n: Count of elements.
' p: Calculated position of quartile.
' j: Element of dataset.
' g: Decimal part of p to be used for interpolation between j and j+1.
'
Public Enum ApQuartileMethod
    [_First] = 1
    
    ' Basic calculation methods.
    
    ' Step. Mendenhall and Sincich method.
    '   SAS #3.
    '   Round up to actual element of dataset.
    '   L:  -Int(-n/4)
    '   H: n-Int(-n/4)
    apMendenhallSincich = 1
    
    ' Average step.
    '   SAS #5, Minitab (%DESCRIBE), GLIM (percentile).    '
    '   Add bias of one on basis of n/4.
    '   L:   CLng((n+2)/2)/2
    '   H: n-Clng((n+2)/2)/2
    '   Note:
    '       Replaces these original formulas that don't return the expected values.
    '   L:   (Int((n+1)/4)+Int(n/4))/2+1
    '   H: n-(Int((n+1)/4)+Int(n/4))/2+1
    apAverage = 2
    
    ' Nearest integer to np.
    '   SAS #2.
    '   Round to nearest integer on basis of n/4.
    '   L:   CLng(n/4)
    '   H: n-CLng(n/4)
    '   Note:
    '       Replaces these original formulas that don't return the expected values.
    '   L:   Int((n+2)/4)
    '   H: n-Int((n+2)/4)
    apNearestInteger = 3
    
    ' Parzen method.
    '   Method 1 with interpolation.
    '   SAS #1.
    '   L: n/4
    '   H: 3n/4
    apParzen = 4
    
    ' Hazen method.
    '   Values midway between method 1 steps.
    '   GLIM (interpolate).
    '   Wikipedia method 3.
    '   Add bias of 2, don't round to actual element of dataset.
    '   L: (n+2)/4
    '   H: 3(n+2)/4-1
    apHazen = 5
    
    ' Weibull method.
    '   SAS #4. Minitab (DECRIBE), SPSS, BMDP, Excel exclusive.
    '   Add bias of 1, don't round to actual element of dataset.
    '   L: (n+1)/4
    '   H: 3(n+1)/4
    apWeibull = 6
    
    ' Freund, J. and Perles, B., Gumbell method.
    '   S-PLUS, R, Excel legacy, Excel inclusive, Star Office Calc.
    '   Add bias of 3, don't round to actual element of dataset.
    '   L: (n+3)/4
    '   H: (3n+1)/4
    apFreundPerlesGumbell = 7
    
    ' Median Position.
    '   Median unbiased.
    '   L: (3n+5)/12
    '   H: (9n+7)/12
    apMedianPosition = 8
    
    ' Bernard and Bos-Levenbach.
    '   L: (n/4)+0.4
    '   H: (3n/4)/+0.6
    '   Note:
    '       Reference claims L to be (n/4)+0.31.
    apBernardBosLevenbach = 9
    
    ' Blom's Plotting Position.
    '   Better approximation when the distribution is normal.
    '   L: (4n+7)/16
    '   H: (12n+9)/16
    apBlom = 10
    
    ' Moore's first method.
    '   Add bias of one half step.
    '   L: (n+0.5)/4
    '   H: n-(n+0.5)/4
    apMooreFirst = 11
    
    ' Moore's second method.
    '   Add bias of one or two steps on basis of (n+1)/4.
    '   L:   (Int((n+1)/4)+Int(n/4))/2+1
    '   H: n-(Int((n+1)/4)+Int(n/4))/2+1
    apMooreSecond = 12
    
    ' John Tukey's method.
    '   Include median from odd dataset in dataset for quartile.
    '   Wikipedia method 2.
    '   L:   (1-Int(-n/2))/2
    '   H: n-(-1-Int(-n/2))/2
    apTukey = 13
    
    ' Moore and McCabe (M & M), variation of John Tukey's method.
    '   TI-83.
    '   Wikipedia method 1.
    '   Exclude median from odd dataset in dataset for quartile.
    '   L:   (Int(n/2)+1)/2
    '   H: n-(Int(n/2)-1)/2
    apTukeyMooreMcCabe = 14
    
    ' Additional variations between Weibull's and Hazen's methods, from
    '   (i-0.000)/(n+1.00)
    ' to
    '   (i-0.500)/(n+0.00)
    
    ' Variation of Weibull.
    '   L: n(n/4-0)/(n+1)
    '   H: n(3n/4-0)/(n+1)
    apWeibullVariation = 15
    
    ' Variation of Blom.
    '   L: n(n/4-3/8)/(n+1/4)
    '   H: n(3n/4-3/8)/(n+1/4)
    apBlomVariation = 16
    
    ' Variation of Tukey.
    '   L: n(n/4-1/3)/(n+1/3)
    '   H: n(3n/4-1/3)/(n+1/3)
    apTukeyVariation = 17
    
    ' Variation of Cunnane.
    '   L: n(n/4-2/5)/(n+1/5)
    '   H: n(3n/4-2/5)/(n+1/5)
    apCunnaneVariation = 18
    
    ' Variation of Gringorten.
    '   L: n(n/4-0.44)/(n+0.12)
    '   H: n(3n/4-0.44)/(n+0.12)
    apGringortenVariation = 19
    
    ' Variation of Hazen.
    '   L: n(n/4-1/2)/n
    '   H: n(3n/4-1/2)/n
    apHazenVariation = 20
    
    [_Last] = 20
End Enum
```

> *If you have comments or corrections to the calculation methods, please post these.*

### Code ###
Code has been tested with both 32-bit and 64-bit *Microsoft Access 2019* and *365*.

### Documentation ###
Full documentation can be found here:

![EE Logo](https://raw.githubusercontent.com/GustavBrock/VBA.Quartiles/master/images/EE%20Logo.png) 

[20 Varieties of Quartiles](https://www.experts-exchange.com/articles/33718/20-Varieties-of-Quartiles.html?preview=%2BptL2cPnHpk%3D)

Included is a Microsoft Access example application.

<hr>

*If you wish to support my work or need extended support or advice, feel free to:*

<p>

[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.Quartiles/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)