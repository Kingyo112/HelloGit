   M S B A M L   `   `   `  ����U  QMahApps.Metro, Version=1.5.0.23, Culture=neutral, PublicKeyToken=f4fb5a3c4d1e5b4fV;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.MetroMahApps.Metro.Controls  5      O KAttendanceManagement, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null;"clr-namespace:AttendanceManagementAttendanceManagement 5      &    AttendanceManagement.StartWindow   5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      PControls;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      4local"clr-namespace:AttendanceManagement     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      -   5          � Title$  StartWindow��5      $	��525��5	      $	��350��5
      (   "MahApps.Metro.Controls.MetroWindow   Flyouts 5      +   %MahApps.Metro.Controls.FlyoutsControl  5   
   .@�5      #   MahApps.Metro.Controls.Flyout  -   5          Name$ flyout��5      $	��525��6.   $��Flyout��6:      Position$ Right=�6J   .��5      �  � RowDefinitions 5      �� 5      $��25 �6(   �� 5      $��25 �6(   �� 5      $��25 �6(   �� 5      $��25 �6(   �� 5      $��25 �6(   �� 5      $��25 �6(   �� 5      $��25 �6(   �� 5      $��25 �6(   �� 5      $��* �6(   �� 5      $��25 �6(    � ColumnDefinitions 5      �� 5      $��50 �6+   �� 5      $	��100 �6+   �� 5      $��* �6+   �� 5       $��50 �6+   .+�5"      �� -   5        $��lblExcelPath_Getudo��5"   9   $��1��6   $��1��6)   $��月度マスタ��6T   �� 5#      -   5        $��txtExcelPath_Getudo��5#   ;   $��1��6   $��2��6+   *   $MahApps.Metro.Controls.TextBoxHelper   ClearTextButton .  �� 
IsReadOnly . 6�   �� 5$      -   5        $��lblExcelPath_Holiday��5$   9   $��3��6   $��1��6)   $��祝日マスタ��6U   �� 5%      -   5        $��txtExcelPath_Holiday��5%   ;   $��3��6   $��2��6+    .  . 6�   �� 5&      -   5        $��lblExcelPath_Attendance��5&   9   $��5��6   $��1��6)   $��勤怠データ��6X   �� 5'      -   5        $��txtExcelPath_Attendance��5'   ;   $��5��6   $��2��6+    .  . 6�   �� 5(      -	   5        $#��lblExcelPath_AttendanceSample��5(   9   $��7��6   $��1��6)   $!��勤怠データサンプル��6^   �� 5)      -
   5        $#��txtExcelPath_AttendanceSample��5)   ;   $��7��6   $��2��6+    .  . 6�   5*      5+      5,      .��5.      �  5/   
   �� 50      $��25 �6   �� 51      $��* �6   �� 52      $��* �6   �� 53      $��25 �6    55   
   �� 56      $��50 �6   �� 57      $��* �6   �� 58      $��* �6   �� 59      $��50 �6   .+�5;   
   !   MahApps.Metro.Controls.Tile  -   5        $��tilAttendance��5;   5   $��1��6   $��1��6%      Title$ 勤怠��6J   	   Count$	 1��6U   
   
TiltFactor$
 2��6_   $
��Auto��6n   $
��Auto��6{   6�     5<   
   -   5        $��
tilMasterM��5<   5   $��1��6   $��2��6%   $ マスタメンテ��6G   $	 2��6V   $
 2��6`   $
��Auto��6o   $
��Auto��6|   6�     5=   
   -   5        $��
tilAnalyze��5=   5   $��2��6   $��1��6%   $ 分析��6G   $	 3��6R   $
 2��6\   $
��Auto��6k   $
��Auto��6x   6�     5>   
   -   5        $��
tilSetting��5>   5   $��2��6   $��2��6%   $ 設定��6G   $	 4��6R   $
 2��6\   $
��Auto��6k   $
��Auto��6x   6�   5?      5@      