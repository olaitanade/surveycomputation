   M S B A M L   `   `   `  ���� H  DMahApps.Metro, Version=1.4.0.8, Culture=neutral, PublicKeyToken=nullV;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.MetroMahApps.Metro.Controls   FiverrExcelForms!   FiverrExcelForms.MainWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     PControls;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro        AccentColorBrush#���   $��1q�(   "MahApps.Metro.Controls.MetroWindow    	GlowBrush#  �    � Title$ SurveyComputation��$	��500��$
��1000��.��� $��7q� � RowDefinitions �� $
��auto ��� $��* ��� $
��auto �.+��� $��1��$��5q�.@��� $��
New Survey��.���� .���� .+��  �� $
��auto ��� $
��auto ��� $
��auto ��� �� $
��auto � � ColumnDefinitions �� $
��auto ��� $��* �.+��� $��2��$��2�� �� Orientation$ 
Horizontal=�$
��Left=�.+��� $��Center=�$��3q�Number of closed loop points�� -   $��closed_loop_pts��$��3q�$��Center=�$��40���� -   $��closed_loop_btn��$��5q�$	��100��.��Create�� $��1��$��2��$ 
Horizontal=�$
��Left=�.+��� $��Center=�$��3q�	StartZ:�� -   $��
StartZ_txt��$��3q�$��Center=�$��80���� $��Center=�$��15,3,3,3q�	StartE:�� -   $��
StartE_txt��$��3q�$��Center=�$��80���� $��Center=�$��15,3,3,3q�	StartN:�� -   $��
StartN_txt��$��3q�$��Center=�$��80���� $��Center=�$��50,3,3,3q�$��17(�&$Total Differential From Start Point:�� -   $��totaldiff_txt��$��Center=�$��10,3,3,3q�$��15(�$��SemiBold"�$��New Times Roman)��� $��0��$
��Left=�$��2��$��17(�$��Italic$�Closed Loop Point:�� $��3��$��2��.��&   System.Windows.Controls.DataGrid  -   $��ClosedDG��$
��Auto=�$
��Auto=�   AutoGenerateColumns .     Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   Header$ Target Point��1  +System.Windows.Controls.DataGridBoundColumn   Binding �� 	 �� Path$	 TargetPoint�
 �� Mode$
 TwoWay=� �� UpdateSourceTrigger$ 	LostFocus=�  $ Stand Point�� �� $	 Standpt�$
 TwoWay=�$ 	LostFocus=�  $ Degree�� �� $	 Degree�$
 TwoWay=�$ 	LostFocus=�  $ Minute�� �� $	 Minute�$
 TwoWay=�$ 	LostFocus=�  $ Second�� �� $	 Second�$
 TwoWay=�$ 	LostFocus=�  $ Distance�� �� $	 Distance�$
 TwoWay=�$ 	LostFocus=�  $ Diff Elev(Delta Z)�� �� $	 DeltaZ�$
 TwoWay=�$ 	LostFocus=�  $ East(X)��   
IsReadOnly .  �� $	 X�$
 TwoWay=�$ PropertyChanged=�  $ North(Y)�� .  �� $	 Y�$
 TwoWay=�$ PropertyChanged=�  $ Elevetion(Z)�� .  �� $	 Z�$
 TwoWay=�$ PropertyChanged=�  $ CW�� .  �� $	 Cw�$
 TwoWay=�$ PropertyChanged=�  $	 CCW�� .  �� $		 Ccw�$
 TwoWay=�$ PropertyChanged=�  $ World Angle�� .  �� $	 WorldA�$
 TwoWay=�$ PropertyChanged=�  $ StartZ�� .  �� $	 StartZ�$
 TwoWay=�$ PropertyChanged=�  $ StartE�� .  �� $	 E�$
 TwoWay=�$ PropertyChanged=�  $ StartN�� .  �� $	 N�$
 TwoWay=�$ PropertyChanged=�  $ Sum Distance�� .  �� $	 Sum_distance�$
 TwoWay=�$ PropertyChanged=�  $ Differential to start(x)�� .  �� $	 Diff_x�$
 TwoWay=�$ PropertyChanged=�  $ Differential to start(y)�� .  �� $	 Diff_y�$
 TwoWay=�$ PropertyChanged=�  $ Differential to start(z)�� .  �� $	 Diff_z�$
 TwoWay=�$ PropertyChanged=��� -   $��closed_loopCompute_btn��$��4��$��2��$��10q�{�.  .��	Compute� $��3,20,3,3q� �� $
��auto ��� $
��auto ��� �� $
��auto ��� $
��auto � �� $
��auto ��� $��* �.+��� $��1��$��2��$ 
Horizontal=�$��Center=�.+��� $��Center=�$��3q�Number of referent points�� -	   $��Referent_pts��$��3q�$��Center=�$��40���� -
   $��Referent_btn��$��5q�$	��100��{�.  .��Create�� $��0��$
��Left=�$��2��$��17(�$��Italic$�Referent Point:  -   $��
referentDG��$��2��$��2�� 4  .System.Windows.Controls.DataGridComboBoxColumn  -   closedpt_combo��$ 	Ref Point��   DisplayMemberPath$ TargetPoint��   SelectedItemBinding �� $	 RefPoint�  $ Target Point�� �� $	 TargetPoint�$
 TwoWay=�$ 	LostFocus=�  $ Degree�� �� $	 Degree�$
 TwoWay=�$ 	LostFocus=�  $ Minute�� �� $	 Minute�$
 TwoWay=�$ 	LostFocus=�  $ Second�� �� $	 Second�$
 TwoWay=�$ 	LostFocus=�  $ Distance�� �� $	 Distance�$
 TwoWay=�$ 	LostFocus=�  $ Diff Elev(Delta Z)�� �� $	 DeltaZ�$
 TwoWay=�$ 	LostFocus=�  $ East(X)�� .  �� $	 X�$
 TwoWay=�$ PropertyChanged=�  $ North(Y)�� .  �� $	 Y�$
 TwoWay=�$ PropertyChanged=�  $ Elevetion(Z)�� .  �� $	 Z�$
 TwoWay=�$ PropertyChanged=�  $ CW�� .  �� $	 Cw�$
 TwoWay=�$ PropertyChanged=�  $	 CCW�� .  �� $		 Ccw�$
 TwoWay=�$ PropertyChanged=�  $ World Angle�� .  �� $	 WorldA�$
 TwoWay=�$ PropertyChanged=�  $ StartE�� .  �� $	 E�$
 TwoWay=�$ PropertyChanged=�  $ StartN�� .  �� $	 N�$
 TwoWay=�$ PropertyChanged=�  $ Differential to start(z)�� .  �� $	 Diff_z�$
 TwoWay=�$ PropertyChanged=��� -   $��ReferentCompute_btn��$��3��$��2��$��10q�{�.  .��	Compute�� -   $��
saveAs_btn��$��4��$��2��$��Right=�$��10q�$��7q�{�.  .��	Save as�� $
��Help��.���� .���� $��10q�$
��Auto=�.+��� $��Center=�$��17(�$��Medium"�$��5q�How To Use The Software9� $��5q�$��5q�$&�� Step 1:  Closed Loop Calculation�� 9� 
IsExpanded . .���� $��13(� �� Padding$ 5q�/-1. Make sure to enter the Starting variables �� .�Startz , Start point in �� .�x ,Start point in �� .�y�� -+2. After which you now enter the number of �� .�Closed Points you will be using.�� TR3. Then click on the create button to create the data grid that you will be using.9� $��5q�$��5q�$��Step 2��.���� $��13(�$ 5q�{y1. Double click on the datagrid cells for your input values. Make sure you enter correct values to yield correct results.�� ��2. After entering the values correctly click on compute and it gives you your answer. Now you can proceed to Referent point calculation9� $��5q�$��5q�$(��"Step 3: Referent Point Calculation��.���� $��13(�$ 5q�861. Enter the number of points used in the calculation.�� 862. Click create and enter all the input values needed.�� �~3. Now you can click compute to give you your answer, after which you can choose to save the entire work as a csv file. Thanks�� $��2��$��Center=�$��5q�Copyright 2016