   M S B A M L   `   `   `  ����S  OFishball_Send_Email_Tool, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullC&clr-namespace:Fishball_Send_Email_ToolFishball_Send_Email_Tool  5   	   )    #Fishball_Send_Email_Tool.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   8local&clr-namespace:Fishball_Send_Email_Tool     5   	     � Title$  鱼蛋邮箱工具 v1.0��5   	   $��399.123��6   $	��525��6.    � WindowStartupLocation$ CenterScreen=�6:    � Icon$ Resoures/logo.ico��6`   .��5	      �  � ColumnDefinitions 5
   
   �� 5      $	��96* �6   �� 5      $
��421* �6   .+�5   
   �� -   5         ��Name$ label1��5      $��SMTP服务器地址：��6    $
��Left=�65   $��	50,17,0,0q�6P   $	��Top=�6c   $��2��6{   �� 5   
   -   5        $ label2��5      $��本人邮箱地址：��6    $
��Left=�62   $��	71,50,0,0q�6M   $	��Top=�6`   $��2��6x   �� 5   
   -   5        $ label3��5      $��本人邮箱密码：��6    $
��Left=�62   $��	70,85,0,0q�6M   $	��Top=�6`   $��2��6x   �� 5   
   -   5        $ label4��5      $��接收人邮件：��6    $
��Left=�61   $��
80,121,0,0q�6L   $	��Top=�6`   $��2��6x   �� 5   
   -   5        $ label5��5      $��邮件标题：��6    $
��Left=�60   $��
92,155,0,0q�6K   $	��Top=�6_   $��2��6w   �� 5   
   -   5        $ label6��5      $��邮件内容：��6    $
��Left=�60   $��
29,188,0,0q�6K   $	��Top=�6_   $��2��6w   �� 5   
   -   5         ��Name$ txt_Smtp��5      $
��Left=�6$   $��23��6?   $��	88,20,0,0q�6K    �� TextWrapping$
 Wrap=�6^   $	��Top=�6r   $	��245��6�   $��1��6�   �� 5   
   -   5        $ txt_MyEmailAddress��5      $
��Left=�6.   $��23��6I   $��	88,54,0,0q�6U   $
 Wrap=�6h   $	��Top=�6|   $	��245��6�   $��1��6�   �� 5   
   -	   5        $ txt_TaEmailAddress��5      $
��Left=�6.   $��23��6I   $��
88,124,0,0q�6U   $
 Wrap=�6i   $	��Top=�6}   $	��245��6�   $��1��6�   �� 5   
   -
   5        $ 	txt_label��5      $
��Left=�6%   $��23��6@   $��
88,157,0,0q�6L   $
 Wrap=�6`   $	��Top=�6t   $	��245��6�   $��1��6�   �� 5   
   -   5        $ txt_content��5      $
��Left=�6'   $��87��6B   $��
50,223,0,0q�6N   $
 Wrap=�6b   $	��Top=�6v   $	��414��6�   $��2��6�   �� 5   
   -   5         ��Name$ btn_sent��5      $��发送��6#   $
��Left=�60   $��130,327,0,0q�6K   $	��Top=�6`   $��64��6x   $��1��6�   G� 5   
   -   5         G�Name$ txt_password��5      $
��Left=�6,   $��	88,88,0,0q�6G   $	��Top=�6Z   $	��245��6r   $��23��6~   $��1��6�   5      5      