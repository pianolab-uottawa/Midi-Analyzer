   M S B A M L   `   `   `  ���� H  DMidi Analyzer, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:Midi_AnalyzerMidi_Analyzer      Midi_Analyzer.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     -localclr-namespace:Midi_Analyzer       � Title$  Midi Analyzer��$	��550��$��899.334��.��� .+�� .+��� -   $��tabController��.@��� $��	Converter��.���� $
��Left=�$	��Top=�.+��� $��Center=�$
��Wrap=�$#��Welcome to the Midi Analyzer!��$	��Top=� �� Left$	 361��
 �� Top$ 15���� $��Center=�$
��Wrap=�$V��PPlease select what file you wish to convert, and a location for the output file.��$	��Top=�$	 243��$ 31���� $
��Left=�$
��Wrap=�$��Source Files:��$	��Top=�$ 99��$	 107���� $��17��$
��Wrap=�$��74��$ 99��$	 285���� $�Excerpt File���� $�:���� �� �� $��Center=�$
��Wrap=�$��Excerpt Picture:��$	��Top=�$ 99��$	 335���� -   $��
excerptBox��$��23�� �� TextWrapping$
 Wrap=�$	��470��$	 230��$	 284������  �� Path$ FlexString1� �� Mode$ TwoWay=� �� UpdateSourceTrigger$ PropertyChanged=�{��� *	IsChecked+ �� ElementName$ 
midiButton���� -   $��imageBox��$��23��$
 Wrap=�$	��470��$	 231��$	 335������ $ FlexString1�$ TwoWay=�$ PropertyChanged=�{��� *	IsChecked+$ 
midiButton���� -   $��Browse��$
��Left=�$	��Top=�$��75��$	 725��$	 285��{��� *	IsChecked+$ 
midiButton���� -   $��Browse��$
��Left=�$	��Top=�$��75��$	 725��$	 338��{��� *	IsChecked+$ 
midiButton���� $��16��$
��Wrap=�$��42��$ 99��$	 385��$��Model:���� -   $��modelBox��$��23��$
 Wrap=�$	��470��$	 231��$	 385������ $ FlexString1�$ TwoWay=�$ PropertyChanged=�{��� *	IsChecked+$ 
midiButton���� -   $��Browse��$
��Left=�$	��Top=�$��75��$	 726��$	 385��{��� *	IsChecked+$ 
midiButton���� $
��Left=�$
��Wrap=�$��Destination Folder:��$	��Top=�$ 99��$	 228���� -   $��
sourcePath��$��92��$	��470��$	 230��$	 107�� �� SelectionMode$ Multiple=��� -	   $��destinationPath��$
��Left=�$��23��$
 Wrap=�$	��Top=�$	��470��$	 230��$	 228������ $ FlexString1�$ TwoWay=�$ PropertyChanged=��� -
   $��Browse��$
��Left=�$	��Top=�$��75��	 J� RenderTransformOrigin$	 
0.178,-0.5&�$	 725��$	 107���� -   $��Browse��$
��Left=�$	��Top=�$��75��$	 725��$	 228���� -   $��Convert��$��Center=�$	��Top=�$��75��$	 326��$	 434���� $
��Left=�$
��Wrap=�$��Source File Type:��$	��Top=�$ 99��$ 69��	� -   $��
midiButton��$
��MIDI��
 	� 	GroupName$
 Group1��$
��Left=�$	��Top=� d� 	IsChecked$
 TrueY�$	 243��$ 73��	� -   $��	csvButton��$	��CSV��$
 Group1��$
��Left=�$	��Top=�$	 361��$ 73���� -   $��Analyze��$
��Left=�$	��Top=�$��75��$	 	0.582,0.4&�$	 531��$	 434��{��� *	IsChecked+$ 
midiButton���� -    ��Name$ errorDetection��$��Error Detection��.���� .+��� -    ��Name$ xlsFileList��$	��198��$	 147��$	 116��$	��603������  �� 
TargetType" ��.�-   �� $
��Wrap=�$ 24��$	 181�� �� TextAlignment$ Center=��� $n�hYour midi files have been converted into a single xls file. During basic error correction, the following���� �� $�files had errors���� $�. ���� $�Please specify in ���� $
��Bold"�$�	Column K ���� $8�2of the sheet corresponding to the file which notes���� �� $� you wish to include (���� $
��Bold"�$�Y���� $�) or exclude ���� $
��Bold"�$	�(N)���� $�. Then, save���� $� the file and press the "���� $��Italic$�$�Finish���� $�" button below.���� �� $	 360��$
��Wrap=�$�� ��$ 59���� -   $��finishButton��$��Finish��$	 408��$	 365��$��75���� -   $ results��$��Results��.���� .+��� $��Center=�$��52��$
��Wrap=�$j��dThe files have been analyzed successfully! You can open the excel file by pressing the button below.��$��16(�$	 194��$ 26��$	��547��$ Center=��� -   $��	Open File��$	��100��$	��243��$	 323��$	 123��