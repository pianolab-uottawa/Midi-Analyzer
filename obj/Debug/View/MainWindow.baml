   M S B A M L   `   `   `  ����H  DMidi Analyzer, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:Midi_AnalyzerMidi_Analyzer  5   	       Midi_Analyzer.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   -localclr-namespace:Midi_Analyzer     5   	     � Title$  Midi Analyzer��5   	   $	��550��6   $��899.334��6,   .��5	      � .+�5
   
   � .+�5      �� -   5        $��tabController��5      .@�5      �� $��	Converter��6   .��5      �� $
��Left=�6   $	��Top=�68   .+�5      �� $��Center=�6$   $
��Wrap=�6A   $#��Welcome to the Midi Analyzer!��6U   $	��Top=�6z    �� Left$	 361��6�   
 �� Top$ 15��6�   �� 5      $��Center=�6$   $
��Wrap=�6A   $V��PPlease select what file you wish to convert, and a location for the output file.��6U   $	��Top=�6�   $	 243��6�   $ 31��6�   �� 5      $
��Left=�6$   $
��Wrap=�6?   $��Source Files:��6S   $	��Top=�6h   $ 99��6�   $	 107��6�   �� 5      $��17��6$   $
��Wrap=�60   $��74��6D   $ 99��6O   $	 285��6`   �� 5      $�Excerpt File��6"   �� 68   $�:��6<   �� 6G   �� 6S   5      �� 5      $��Center=�6$   $
��Wrap=�6A   $��Excerpt Picture:��6U   $	��Top=�6m   $ 99��6�   $	 335��6�   6�   �� 5      -   5        $��
excerptBox��5   t   $��23��5   "    �� TextWrapping$
 Wrap=�6.   $	��470��6B   $	 230��6N   $	 284��6`   ��5   "   ��  �� Path$ FlexString1� �� Mode$ TwoWay=� �� UpdateSourceTrigger$ PropertyChanged=�{�5   "   �� *	IsChecked+ �� ElementName$ 
midiButton���� 5      -   5        $��imageBox��5   t   $��23��5   "   $
 Wrap=�6.   $	��470��6B   $	 231��6N   $	 335��6`   ��5   "   �� $ FlexString1�$ TwoWay=�$ PropertyChanged=�{�5   "   �� *	IsChecked+$ 
midiButton���� 5      -   5        $��Browse��5   !   $
��Left=�62   $	��Top=�6M   $��75��5      $	 725��6=   $	 285��6O   {�5   !   �� *	IsChecked+$ 
midiButton���� 5      -   5        $��Browse��5   !   $
��Left=�62   $	��Top=�6M   $��75��5      $	 725��6?   $	 338��6Q   {�5    !   �� *	IsChecked+$ 
midiButton���� 5!      $��16��6$   $
��Wrap=�60   $��42��6D   $ 99��6O   $	 385��6`   $��Model:��6q   �� 5"      -   5        $��modelBox��5"   t   $��23��5#   "   $
 Wrap=�6.   $	��470��6B   $	 231��6N   $	 385��6`   ��5"   "   �� $ FlexString1�$ TwoWay=�$ PropertyChanged=�{�5$   "   �� *	IsChecked+$ 
midiButton���� 5%      -   5        $��Browse��5%   !   $
��Left=�62   $	��Top=�6M   $��75��5&      $	 726��6>   $	 385��6P   {�5'   !   �� *	IsChecked+$ 
midiButton���� 5(      $
��Left=�6$   $
��Wrap=�6?   $��Destination Folder:��6S   $	��Top=�6n   $ 99��6�   $	 228��6�   �� 5)      -   5        $��
sourcePath��5)   :   $��92��6"   $	��470��6.   $	 230��6L   $	 107��6^    �� SelectionMode$ Multiple=�6�   �� 5*      -	   5        $��destinationPath��5+      $
��Left=�64   $��23��6O   $
 Wrap=�5,      $	��Top=�61   $	��470��6I   $	 230��6U   $	 228��6g   ��5*   "   �� $ FlexString1�$ TwoWay=�$ PropertyChanged=�5,   g   �� 5-      -
   5        $��Browse��5-   !   $
��Left=�62   $	��Top=�6M   $��75��5.      	 J� RenderTransformOrigin$	 
0.178,-0.5&�6F   $	 725��6i   $	 107��6{   �� 5/      -   5        $��Browse��5/   !   $
��Left=�62   $	��Top=�6M   $��75��50      $	 725��6@   $	 228��6R   �� 51      -   5        $��Convert��51   !   $��Center=�63   $	��Top=�6P   $��75��52      $	 326��6<   $	 434��6N   �� 53      $
��Left=�6$   $
��Wrap=�6?   $��Source File Type:��6S   $	��Top=�6l   $ 99��6�   $ 69��6�   	� 54      -   5        $��
midiButton��54   &   $
��MIDI��68   
 	� 	GroupName$
 Group1��6G   $
��Left=�6Z   $	��Top=�6u    d� 	IsChecked$
 TrueY�6�   $	 243��6�   $ 73��6�   	� 55      -   5        $��	csvButton��55   &   $	��CSV��67   $
 Group1��6E   $
��Left=�6X   $	��Top=�6s   $	 361��6�   $ 73��6�   �� 56      -   5        $��Analyze��56   !   $
��Left=�63   $	��Top=�6N   $��75��6f   $	 	0.582,0.4&�6q   $	 531��57   m   $	 434��6   {�65   �� *	IsChecked+$ 
midiButton��6   58      59      �� 5:      -   5         ��Name$ errorDetection��5:      $��Error Detection��62   .��5;      �� .+�5<      �� -   5         ��Name$ xlsFileList��5<   "   $	��198��67   $	 147��6D   $	 116��6V   $	��603��6g   ��5=      �� 5>   "    �� 
TargetType" ��6(   .�5?   &   -   5        5@   #   5B      �� 5C      $
��Wrap=�6$   $ 24��68   $	 181��6H    �� TextAlignment$ Center=�6Z   �� 6r   $n�hYour midi files have been converted into a single xls file. During basic error correction, the following��6v   �� 6�   �� 6�   $�files had errors��6�   �� 6  $�. ��6  �� 6"  $�Please specify in ��6&  �� 6B  $
��Bold"�6F  $�	Column K ��6X  �� 6k  $8�2of the sheet corresponding to the file which notes��6o  �� 6�  �� 6�  $� you wish to include (��6�  �� 6�  $
��Bold"�6�  $�Y��6�  �� 6�  $�) or exclude ��6   �� 6  $
��Bold"�6  $	�(N)��6-  �� 6:  $�. Then, save��6>  �� 6T  $� the file and press the "��6X  �� 6�  $��Italic$�6�  $�Finish��6�  �� 6�  $�" button below.��6�  �� 6�  6�  �� 5D      $	 360��6$   $
��Wrap=�66   $�� ��6J   $ 59��6R   �� 5E      -   5        $��finishButton��5E   `   $��Finish��6!   $	 408��62   $	 365��6D   $��75��6U   6t   5F      5G      �� 5H      -   5        $ results��5H      $��Results��6+   .��5I      �� .+�5J      �� $��Center=�6$   $��52��6A   $
��Wrap=�6M   $j��dThe files have been analyzed successfully! You can open the excel file by pressing the button below.��6a   $��16(�6�   $	 194��6�   $ 26��6�   $	��547��6�   $ Center=�6	  �� 5K      -   5        $��	Open File��5K   !   $	��100��65   $	��243��6B   $	 323��6N   $	 123��6`   6q   5L      5M      5N      5O      5P      5Q      