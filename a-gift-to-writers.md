#1 

 

 ****** A Gift to Writers ****** byTx_Tall_Tales© ======================================= 

 It's/its, your's/yours, to/too/two, etc. 

 I've learned a lot writing for Literotica the last dozen years. I think it's high time I paid back a little for all the help I've received during that time. Here's a simple tool I use regularly to assist with some of the most common errors I fell prey to in my first decade of writing. Using it's for its, you're for your, misusing homonyms, which I did far too often. I notice it's a common problem, and hope someone else finds it useful. 

 It's well short of the 750 word minimum, and I'd hate to clutter something so simple with an extra 200 words of fluff. 

 ======================================= 

 The following Word macro, when run, will highlight in RED all the common homonym's that are listed. If you, like me, occasionally type the wrong one, it'll allow you to go back and identify each potentially incorrect word, so you can determine if you got it right or not. 

 It's a very basic macro, and doesn't deal with footnotes or headers or anything of that type. Simple body text is all it covers. The original was written by a friend of mine, and I've modified it for my own needs. Of course, you're encouraged to do the same. 

 To use the Macros, simply go to Macros -> View Macros and press create. You'll have to enter the name of the macro (hilite_HOMONYMS). Copy the macros below into the open file and save. When you View Macros again, you'll have two new ones. The first, hilite_HOMONYMS will highlight all the problem words in red. The second one, unhilite_ALL, will clear all highlighting. 

 Once you've save this to your normal.dot, these macros should be available each time you open word. 

 --- Macro starts below here --- 

 Sub hilite_HOMONYMS() ' ' hilite_HOMONYMS Macro ' Macro created 5/27/2013 by Tx Tall Tales ' 

 Dim varWordList(45) As String 

 varWordList(0) = "accept" varWordList(1) = "except" varWordList(2) = "already" varWordList(3) = "all ready" varWordList(4) = "all together" varWordList(5) = "altogether" varWordList(6) = "altar" varWordList(7) = "alter" varWordList(8) = "ascent" varWordList(9) = "assent" 

 varWordList(10) = "bare" varWordList(11) = "bear" varWordList(12) = "brake" varWordList(13) = "break" varWordList(14) = "capital" varWordList(15) = "capitol" varWordList(16) = "conscience" varWordList(17) = "concious" varWordList(18) = "desert" varWordList(19) = "dessert" 

 varWordList(20) = "emigrate" varWordList(21) = "immigrate" varWordList(22) = "its" varwordList(23) = "it's" varWordList(24) = "lead" varWordList(25) = "led" varWordList(26) = "loose" varWordList(27) = "lose" varWordList(28) = "passed" varWordList(29) = "past" 

 varWordList(30) = "principal" varWordList(31) = "principle" varWordList(32) = "their" varWordList(33) = "there" varWordList(34) = "they're" varWordList(35) = "to" varWordList(36) = "too" varWordList(37) = "two" varWordList(38) = "weather" varWordList(39) = "whether" 

 varWordList(40) = "your" varWordList(41) = "you're" varWordList(43) = "end" varWordList(44) = "end" varWordList(45) = "end" 

 counter = 0 

 Do 

 With ActiveDocument.Content.Find .ClearFormatting .Replacement.ClearFormatting .Replacement.Font.Color = wdColorRed .MatchWholeWord = True .MatchCase = False ' .MatchWildcards = False ' .MatchSoundsLike = False ' .MatchAllWordForms = False .Execute FindText:=varWordList(counter), _ ReplaceWith:=varWordList(counter), Replace:=wdReplaceAll End With 

 counter = counter + 1 

 Loop Until "end" = varWordList(counter) 

 End Sub 

 

 

 Sub unhilite() ' ' unhilite Macro ' Macro created 5/27/2013 by Tx Tall Tales ' 

 With ActiveDocument.Content.Find .ClearFormatting .Font.Color = wdColorRed With .Replacement .ClearFormatting .Font.Color = wdColorBlack End With .Execute FindText:="", ReplaceWith:="", _ Format:=True, Replace:=wdReplaceAll End With 

 End Sub Report_Story 