FasdUAS 1.101.10   ��   ��    k             i         I      �� 	���� 0 opentemplate OpenTemplate 	  
�� 
 o      ���� $0 initialdirectory initialDirectory��  ��    Q          k           r        I   ���� 
�� .sysostdfalis    ��� null��    ��  
�� 
prmp  m       �   J L o a d   t h e   S p e a k i n g   E v a l u a t i o n   T e m p l a t e  ��  
�� 
ftyp  J    
    ��  m       �    d o c x��    �� ��
�� 
dflc  o    ���� $0 initialdirectory initialDirectory��    o      ���� (0 evaluationtemplate evaluationTemplate   ��  L         o    ���� (0 evaulationtemplate evaulationTemplate��    R      ������
�� .ascrerr ****      � ****��  ��    L     ! ! m     " " � # #     $ % $ l     ��������  ��  ��   %  & ' & i     ( ) ( I      �� *����  0 selectsavepath SelectSavePath *  +�� + o      ���� $0 initialdirectory initialDirectory��  ��   ) Q      , - . , k     / /  0 1 0 r     2 3 2 I   ���� 4
�� .sysostflalis    ��� null��   4 �� 5 6
�� 
prmp 5 m     7 7 � 8 8 Z S e l e c t   W h e r e   t o   S a v e   t h e   S p e a k i n g   E v a u l a t i o n s 6 �� 9��
�� 
dflc 9 o    ���� $0 initialdirectory initialDirectory��   3 o      ���� &0 destinationfolder destinationFolder 1  :�� : L     ; ; o    ���� &0 destinationfolder destinationFolder��   - R      ������
�� .ascrerr ****      � ****��  ��   . L     < < m     = = � > >   '  ? @ ? l     ��������  ��  ��   @  A B A i     C D C I      �� E���� "0 getmacosversion GetMacOSVersion E  F�� F o      ���� 0 paramstring paramString��  ��   D Q      G H�� G k     I I  J K J r    
 L M L I   �� N��
�� .sysoexecTEXT���     TEXT N m     O O � P P . s w _ v e r s   - p r o d u c t V e r s i o n��   M o      ���� 0 	osversion 	osVersion K  Q�� Q L     R R o    ���� 0 	osversion 	osVersion��   H R      ������
�� .ascrerr ****      � ****��  ��  ��   B  S T S l     ��������  ��  ��   T  U V U i     W X W I      �� Y���� 00 getscriptversionnumber GetScriptVersionNumber Y  Z�� Z o      ���� 0 paramstring paramString��  ��   X L      [ [ m     ���� 4�� V  \ ] \ l     ��������  ��  ��   ]  ^ _ ^ i     ` a ` I      �� b���� 00 getlatestscriptversion GetLatestScriptVersion b  c�� c o      ���� *0 onlinescriptversion onlineScriptVersion��  ��   a k     ^ d d  e f e Z      g h�� i g ?      j k j o     ���� *0 onlinescriptversion onlineScriptVersion k I    �� l���� 00 getscriptversionnumber GetScriptVersionNumber l  m�� m m     n n � o o  n o P a r a m��  ��   h r     p q p I    �� r���� :0 downloadlatestscriptversion DownloadLatestScriptVersion r  s�� s m     t t � u u  n o P a r a m��  ��   q o      ���� (0 downloadsuccessful downloadSuccessful��   i L     v v m    ��
�� boovfals f  w x w l   ��������  ��  ��   x  y z y r    $ { | { b    " } ~ } n       �  1     ��
�� 
psxp � l    ����� � I   �� ���
�� .earsffdralis        afdr � m    ��
�� afdrcusr��  ��  ��   ~ m     ! � � � � � j / L i b r a r y / A p p l i c a t i o n   S c r i p t s / c o m . m i c r o s o f t . P o w e r p o i n t | o      ���� 0 scriptfolder scriptFolder z  � � � r   % * � � � b   % ( � � � o   % &���� 0 scriptfolder scriptFolder � m   & ' � � � � � ( / A n g r y B i r d s T e m p . s c p t � o      ���� .0 latestversionfilepath latestVersionFilePath �  � � � r   + 0 � � � b   + . � � � o   + ,���� 0 scriptfolder scriptFolder � m   , - � � � � �   / A n g r y B i r d s . s c p t � o      ����  0 targetfilepath targetFilePath �  � � � l  1 1��������  ��  ��   �  ��� � Q   1 ^ � � � � k   4 J � �  � � � I  4 G�� ���
�� .sysoexecTEXT���     TEXT � b   4 C � � � b   4 = � � � b   4 ; � � � m   4 5 � � � � �  m v   � n   5 : � � � 1   8 :��
�� 
strq � n   5 8 � � � 1   6 8��
�� 
psxp � o   5 6���� .0 latestversionfilepath latestVersionFilePath � m   ; < � � � � �    � n   = B � � � 1   @ B��
�� 
strq � n   = @ � � � 1   > @��
�� 
psxp � o   = >����  0 targetfilepath targetFilePath��   �  ��� � L   H J � � m   H I��
�� boovtrue��   � R      �� ���
�� .ascrerr ****      � **** � o      ���� 0 errmsg errMsg��   � k   R ^ � �  � � � I  R [�� ���
�� .sysodlogaskr        TEXT � b   R W � � � m   R U � � � � � @ E r r o r   u p d a t i n g   A n g r y B i r d s . s c p t :   � o   U V���� 0 errmsg errMsg��   �  ��� � L   \ ^ � � m   \ ]��
�� boovfals��  ��   _  � � � l     ��������  ��  ��   �  � � � i     � � � I      �� ����� :0 downloadlatestscriptversion DownloadLatestScriptVersion �  ��� � o      ���� 0 paramstring paramString��  ��   � Q     5 � � � � k    % � �  � � � r     � � � m     � � � � � � h t t p s : / / g i t h u b . c o m / p a p e r c u t t e r 0 3 2 4 / A n g r y B i r d s T r i v i a - A d d i t i o n a l F i l e s / r a w / m a i n / A n g r y B i r d s . s c p t � o      ���� 0 githubrawurl githubRawURL �  � � � r     � � � b     � � � n     � � � 1    ��
�� 
psxp � l    ����� � I   �� ���
�� .earsffdralis        afdr � m    �
� afdrcusr��  ��  ��   � m     � � � � � � / L i b r a r y / A p p l i c a t i o n   S c r i p t s / c o m . m i c r o s o f t . P o w e r p o i n t / A n g r y B i r d s T e m p . s c p t � o      �~�~ 0 downloadpath downloadPath �  � � � I   "�} ��|
�} .sysoexecTEXT���     TEXT � b     � � � b     � � � b     � � � m     � � � � �  c u r l   - L   - o   � n     � � � 1    �{
�{ 
strq � o    �z�z 0 downloadpath downloadPath � m     � � � � �    � n     � � � 1    �y
�y 
strq � o    �x�x 0 githubrawurl githubRawURL�|   �  ��w � L   # % � � m   # $�v
�v boovtrue�w   � R      �u�t�s
�u .ascrerr ****      � ****�t  �s   � k   - 5 � �  � � � I  - 2�r ��q
�r .sysodlogaskr        TEXT � m   - . � � � � � f E r r o r   d o w n l o a d i n g   l a t e s t   v e r s i o n   o f   A n g r y B i r d s . s c p t�q   �  ��p � L   3 5 � � m   3 4�o
�o boovfals�p   �  � � � l     �n�m�l�n  �m  �l   �  ��k � i     � � � I      �j ��i�j 0 downloadfile DownloadFile �  ��h � o      �g�g 0 paramstring paramString�h  �i   � k     > � �  �  � r      I      �f�e�f 0 splitstring SplitString  o    �d�d 0 paramstring paramString �c m     �  ,�c  �e   J      		 

 o      �b�b 0 fileurl fileURL �a o      �`�` 0 savepath savePath�a    �_ Q    > k    ,  I   )�^�]
�^ .sysoexecTEXT���     TEXT b    % b    ! b     m     �  c u r l   - L   - o   l   �\�[ n     1    �Z
�Z 
strq o    �Y�Y 0 savepath savePath�\  �[   m        �!!    n   ! $"#" 1   " $�X
�X 
strq# o   ! "�W�W 0 fileurl fileURL�]   $�V$ L   * ,%% m   * +�U
�U boovtrue�V   R      �T�S�R
�T .ascrerr ****      � ****�S  �R   k   4 >&& '(' I  4 ;�Q)�P
�Q .sysodlogaskr        TEXT) b   4 7*+* m   4 5,, �-- 0 E r r o r   d o w n l o a d i n g   f i l e :  + o   5 6�O�O 0 fileurl fileURL�P  ( .�N. L   < >// m   < =�M
�M boovfals�N  �_  �k       	�L01234567�L  0 �K�J�I�H�G�F�E�K 0 opentemplate OpenTemplate�J  0 selectsavepath SelectSavePath�I "0 getmacosversion GetMacOSVersion�H 00 getscriptversionnumber GetScriptVersionNumber�G 00 getlatestscriptversion GetLatestScriptVersion�F :0 downloadlatestscriptversion DownloadLatestScriptVersion�E 0 downloadfile DownloadFile1 �D �C�B89�A�D 0 opentemplate OpenTemplate�C �@:�@ :  �?�? $0 initialdirectory initialDirectory�B  8 �>�=�<�> $0 initialdirectory initialDirectory�= (0 evaluationtemplate evaluationTemplate�< (0 evaulationtemplate evaulationTemplate9 
�; �: �9�8�7�6�5 "
�; 
prmp
�: 
ftyp
�9 
dflc�8 
�7 .sysostdfalis    ��� null�6  �5  �A   *����kv�� E�O�W 	X  �2 �4 )�3�2;<�1�4  0 selectsavepath SelectSavePath�3 �0=�0 =  �/�/ $0 initialdirectory initialDirectory�2  ; �.�-�. $0 initialdirectory initialDirectory�- &0 destinationfolder destinationFolder< �, 7�+�*�)�(�' =
�, 
prmp
�+ 
dflc�* 
�) .sysostflalis    ��� null�(  �'  �1  *���� E�O�W 	X  �3 �& D�%�$>?�#�& "0 getmacosversion GetMacOSVersion�% �"@�" @  �!�! 0 paramstring paramString�$  > � ��  0 paramstring paramString� 0 	osversion 	osVersion?  O���
� .sysoexecTEXT���     TEXT�  �  �#  �j E�O�W X  h4 � X��AB�� 00 getscriptversionnumber GetScriptVersionNumber� �C� C  �� 0 paramstring paramString�  A �� 0 paramstring paramStringB �� 4��� �5 � a��DE�� 00 getlatestscriptversion GetLatestScriptVersion� �F� F  �� *0 onlinescriptversion onlineScriptVersion�  D ����
�	�� *0 onlinescriptversion onlineScriptVersion� (0 downloadsuccessful downloadSuccessful� 0 scriptfolder scriptFolder�
 .0 latestversionfilepath latestVersionFilePath�	  0 targetfilepath targetFilePath� 0 errmsg errMsgE  n� t���� � � � �� ��� �� ���� 00 getscriptversionnumber GetScriptVersionNumber� :0 downloadlatestscriptversion DownloadLatestScriptVersion
� afdrcusr
� .earsffdralis        afdr
� 
psxp
� 
strq
� .sysoexecTEXT���     TEXT�  0 errmsg errMsg��  
�� .sysodlogaskr        TEXT� _�*�k+  *�k+ E�Y fO�j �,�%E�O��%E�O��%E�O ��,�,%�%��,�,%j OeW X  a �%j Of6 �� �����GH���� :0 downloadlatestscriptversion DownloadLatestScriptVersion�� ��I�� I  ���� 0 paramstring paramString��  G �������� 0 paramstring paramString�� 0 githubrawurl githubRawURL�� 0 downloadpath downloadPathH  ������� � ��� ������� ���
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  
�� .sysodlogaskr        TEXT�� 6 '�E�O�j �,�%E�O��,%�%��,%j OeW X 	 
�j Of7 �� �����JK���� 0 downloadfile DownloadFile�� ��L�� L  ���� 0 paramstring paramString��  J �������� 0 paramstring paramString�� 0 fileurl fileURL�� 0 savepath savePathK ������ ������,���� 0 splitstring SplitString
�� 
cobj
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  
�� .sysodlogaskr        TEXT�� ?*��l+ E[�k/E�Z[�l/E�ZO ��,%�%��,%j OeW X  �%j 
Ofascr  ��ޭ