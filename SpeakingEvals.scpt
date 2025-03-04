FasdUAS 1.101.10   ��   ��    k             l      ��  ��    � |
Helper Scripts for the DYB Speaking Evaluations Excel spreadsheet

Version: 1.3.2
Build:   20250306
Warren Feltmate
� 2025
     � 	 	 � 
 H e l p e r   S c r i p t s   f o r   t h e   D Y B   S p e a k i n g   E v a l u a t i o n s   E x c e l   s p r e a d s h e e t 
 
 V e r s i o n :   1 . 3 . 2 
 B u i l d :       2 0 2 5 0 3 0 6 
 W a r r e n   F e l t m a t e 
 �   2 0 2 5 
   
  
 l     ��������  ��  ��        l     ��  ��      Environment Variables     �   ,   E n v i r o n m e n t   V a r i a b l e s      l     ��������  ��  ��        i         I      �� ���� 00 getscriptversionnumber GetScriptVersionNumber   ��  o      ���� 0 paramstring paramString��  ��    k            l     ��  ��    ? 9- Use build number to determine if an update is available     �   r -   U s e   b u i l d   n u m b e r   t o   d e t e r m i n e   i f   a n   u p d a t e   i s   a v a i l a b l e   ��  L          m     ���� 4����     ! " ! l     ��������  ��  ��   "  # $ # i     % & % I      �� '���� "0 getmacosversion GetMacOSVersion '  (�� ( o      ���� 0 paramstring paramString��  ��   & k      ) )  * + * l     �� , -��   , ` Z Not currently used, but could be helpful if there are issues with older versions of MacOS    - � . . �   N o t   c u r r e n t l y   u s e d ,   b u t   c o u l d   b e   h e l p f u l   i f   t h e r e   a r e   i s s u e s   w i t h   o l d e r   v e r s i o n s   o f   M a c O S +  /�� / Q      0 1�� 0 k     2 2  3 4 3 r    
 5 6 5 I   �� 7��
�� .sysoexecTEXT���     TEXT 7 m     8 8 � 9 9 . s w _ v e r s   - p r o d u c t V e r s i o n��   6 o      ���� 0 	osversion 	osVersion 4  :�� : L     ; ; o    ���� 0 	osversion 	osVersion��   1 R      ������
�� .ascrerr ****      � ****��  ��  ��  ��   $  < = < l     ��������  ��  ��   =  > ? > i     @ A @ I      �� B���� 80 checkaccessibilitysettings CheckAccessibilitySettings B  C�� C o      ���� 0 
apptocheck 
appToCheck��  ��   A k     5 D D  E F E l     �� G H��   G � { Not used yet, but might be in the future as a way to validate and correct invalid entries, such as with a student's grades    H � I I �   N o t   u s e d   y e t ,   b u t   m i g h t   b e   i n   t h e   f u t u r e   a s   a   w a y   t o   v a l i d a t e   a n d   c o r r e c t   i n v a l i d   e n t r i e s ,   s u c h   a s   w i t h   a   s t u d e n t ' s   g r a d e s F  J�� J Q     5 K L M K O    + N O N k    * P P  Q R Q l   �� S T��   S O I Checks if Accessibility features are enabled for the checked application    T � U U �   C h e c k s   i f   A c c e s s i b i l i t y   f e a t u r e s   a r e   e n a b l e d   f o r   t h e   c h e c k e d   a p p l i c a t i o n R  V W V r    ' X Y X F    % Z [ Z l    \���� \ E    ] ^ ] l    _���� _ 6   ` a ` n     b c b 1   
 ��
�� 
pnam c 2   
��
�� 
prcs a =    d e d 1    ��
�� 
pvis e m    ��
�� boovtrue��  ��   ^ o    ���� 0 
apptocheck 
appToCheck��  ��   [ l   # f���� f n    # g h g 1     "��
�� 
enaB h n      i j i 2    ��
�� 
uiel j 4    �� k
�� 
pcap k o    ���� 0 
apptocheck 
appToCheck��  ��   Y o      ���� ,0 accessibilityenabled accessibilityEnabled W  l�� l L   ( * m m o   ( )���� ,0 accessibilityenabled accessibilityEnabled��   O m     n n�                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��   L R      ������
�� .ascrerr ****      � ****��  ��   M L   3 5 o o m   3 4��
�� boovfals��   ?  p q p l     ��������  ��  ��   q  r s r l     �� t u��   t   Parameter Manipulation    u � v v .   P a r a m e t e r   M a n i p u l a t i o n s  w x w l     ��������  ��  ��   x  y z y i     { | { I      �� }���� 0 splitstring SplitString }  ~  ~ o      ���� &0 passedparamstring passedParamString   ��� � o      ���� (0 parameterseparator parameterSeparator��  ��   | k      � �  � � � l     �� � ���   � d ^ Excel can only pass on parameter to this file. This makes it possible to split one into many.    � � � � �   E x c e l   c a n   o n l y   p a s s   o n   p a r a m e t e r   t o   t h i s   f i l e .   T h i s   m a k e s   i t   p o s s i b l e   t o   s p l i t   o n e   i n t o   m a n y . �  � � � O      � � � k     � �  � � � r    	 � � � 1    ��
�� 
txdl � o      ���� 00 oldtextitemsdelimiters oldTextItemsDelimiters �  � � � r   
  � � � o   
 ���� (0 parameterseparator parameterSeparator � 1    ��
�� 
txdl �  � � � r     � � � n     � � � 2   ��
�� 
citm � o    ���� &0 passedparamstring passedParamString � o      ���� *0 separatedparameters separatedParameters �  ��� � r     � � � o    ���� 00 oldtextitemsdelimiters oldTextItemsDelimiters � 1    ��
�� 
txdl��   � 1     ��
�� 
ascr �  ��� � L     � � o    ���� *0 separatedparameters separatedParameters��   z  � � � l     ��������  ��  ��   �  � � � l     �� � ���   �    Application Manipulations    � � � � 4   A p p l i c a t i o n   M a n i p u l a t i o n s �  � � � l     ��������  ��  ��   �  � � � i     � � � I      �� ����� "0 loadapplication LoadApplication �  ��� � o      ���� 0 appname appName��  ��   � k     ) � �  � � � l     �� � ���   � < 6 A simple function to tell the needed program to open.    � � � � l   A   s i m p l e   f u n c t i o n   t o   t e l l   t h e   n e e d e d   p r o g r a m   t o   o p e n . �  ��� � Q     ) � � � � k     � �  � � � O    � � � I  
 ������
�� .miscactvnull��� ��� null��  ��   � 4    �� �
�� 
capp � o    ���� 0 appname appName �  ��� � L     � � m     � � � � �  ��   � R      �� � �
�� .ascrerr ****      � **** � o      ���� 0 errmsg errMsg � �� ���
�� 
errn � o      ���� 0 errnum errNum��   � L    ) � � b    ( � � � b    & � � � b    $ � � � b    " � � � b      � � � b     � � � m     � � � � �  E r r o r   l o a d i n g � 1    ��
�� 
spac � o    ���� 0 appname appName � m     ! � � � � �  :   � o   " #���� 0 errnum errNum � m   $ % � � � � �    -   � o   & '���� 0 errmsg errMsg��   �  � � � l     ��������  ��  ��   �  � � � i     � � � I      �� ����� 0 isapploaded IsAppLoaded �  ��� � o      ���� 0 appname appName��  ��   � k     : � �  � � � l     �� � ���   � N H This lets Excel check that the other program is open before continuing.    � � � � �   T h i s   l e t s   E x c e l   c h e c k   t h a t   t h e   o t h e r   p r o g r a m   i s   o p e n   b e f o r e   c o n t i n u i n g . �  �� � Q     : � � � � k    & � �  � � � O    # � � � Z    " � ��~ � � E     � � � l    ��}�| � n     � � � 1   
 �{
�{ 
pnam � 2    
�z
�z 
prcs�}  �|   � o    �y�y 0 appname appName � r     � � � b     � � � b     � � � o    �x�x 0 appname appName � 1    �w
�w 
spac � m     � � � � �  i s   n o w   r u n n i n g . � o      �v�v 0 
loadresult 
loadResult�~   � r    " �  � b      b     m     �  E r r o r   o p e n i n g 1    �u
�u 
spac o    �t�t 0 appname appName  o      �s�s 0 
loadresult 
loadResult � m    �                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��   � �r L   $ &		 o   $ %�q�q 0 
loadresult 
loadResult�r   � R      �p

�p .ascrerr ****      � ****
 o      �o�o 0 errmsg errMsg �n�m
�n 
errn o      �l�l 0 errnum errNum�m   � L   . : b   . 9 b   . 7 b   . 5 b   . 3 b   . 1 m   . / �  E r r o r   l o a d i n g   o   / 0�k�k 0 appname appName m   1 2 �  :   o   3 4�j�j 0 errnum errNum m   5 6 �    -   o   7 8�i�i 0 errmsg errMsg�   �  l     �h�g�f�h  �g  �f    !  i    "#" I      �e$�d�e "0 closepowerpoint ClosePowerPoint$ %�c% o      �b�b 0 paramstring paramString�c  �d  # k     3&& '(' l     �a)*�a  ) { u This will completely close MS PowerPoint, even from the Dock. This reduces the chances of errors on subsequent runs.   * �++ �   T h i s   w i l l   c o m p l e t e l y   c l o s e   M S   P o w e r P o i n t ,   e v e n   f r o m   t h e   D o c k .   T h i s   r e d u c e s   t h e   c h a n c e s   o f   e r r o r s   o n   s u b s e q u e n t   r u n s .( ,�`, Q     3-./- O    )010 k    (22 343 Z    %56�_75 E    898 l   :�^�]: n    ;<; 1   
 �\
�\ 
pnam< 2    
�[
�[ 
prcs�^  �]  9 m    == �>> ( M i c r o s o f t   P o w e r P o i n t6 k    ?? @A@ O   BCB I   �Z�Y�X
�Z .aevtquitnull��� ��� null�Y  �X  C m    DD�                                                                                  PPT3  alis    L  macOS                      ��'�BD ����Microsoft PowerPoint.app                                       �����Ώ�        ����  
 cu             Applications  (/:Applications:Microsoft PowerPoint.app/  2  M i c r o s o f t   P o w e r P o i n t . a p p    m a c O S  %Applications/Microsoft PowerPoint.app   / ��  A E�WE r    FGF m    HH �II P P o w e r P o i n t   h a s   s u c c e s s f u l l y   b e e n   c l o s e d .G o      �V�V 0 closeresult closeResult�W  �_  7 r   " %JKJ m   " #LL �MM H P o w e r P o i n t   i s   n o t   c u r r e n t l y   r u n n i n g .K o      �U�U 0 closeresult closeResult4 N�TN L   & (OO o   & '�S�S 0 closeresult closeResult�T  1 m    PP�                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  . R      �R�Q�P
�R .ascrerr ****      � ****�Q  �P  / L   1 3QQ m   1 2RR �SS \ T h e r e   w a s   a n   e r r o r   t r y i n g   t o   c l o s e   P o w e r P o i n t .�`  ! TUT l     �O�N�M�O  �N  �M  U VWV i    XYX I      �LZ�K�L 0 	closeword 	CloseWordZ [�J[ o      �I�I 0 paramstring paramString�J  �K  Y k     3\\ ]^] l     �H_`�H  _ u o This will completely close MS Word, even from the Dock. This reduces the chances of errors on subsequent runs.   ` �aa �   T h i s   w i l l   c o m p l e t e l y   c l o s e   M S   W o r d ,   e v e n   f r o m   t h e   D o c k .   T h i s   r e d u c e s   t h e   c h a n c e s   o f   e r r o r s   o n   s u b s e q u e n t   r u n s .^ b�Gb Q     3cdec O    )fgf k    (hh iji Z    %kl�Fmk E    non l   p�E�Dp n    qrq 1   
 �C
�C 
pnamr 2    
�B
�B 
prcs�E  �D  o m    ss �tt  M i c r o s o f t   W o r dl k    uu vwv O   xyx I   �A�@�?
�A .aevtquitnull��� ��� null�@  �?  y m    zz�                                                                                  MSWD  alis    4  macOS                      ��'�BD ����Microsoft Word.app                                             �����Ώ�        ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    m a c O S  Applications/Microsoft Word.app   / ��  w {�>{ r    |}| m    ~~ � D W o r d   h a s   s u c c e s s f u l l y   b e e n   c l o s e d .} o      �=�= 0 closeresult closeResult�>  �F  m r   " %��� m   " #�� ��� < W o r d   i s   n o t   c u r r e n t l y   r u n n i n g .� o      �<�< 0 closeresult closeResultj ��;� L   & (�� o   & '�:�: 0 closeresult closeResult�;  g m    ���                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  d R      �9�8�7
�9 .ascrerr ****      � ****�8  �7  e L   1 3�� m   1 2�� ��� P T h e r e   w a s   a n   e r r o r   t r y i n g   t o   c l o s e   W o r d .�G  W ��� l     �6�5�4�6  �5  �4  � ��� l     �3���3  �   File Manipulation   � ��� $   F i l e   M a n i p u l a t i o n� ��� l     �2�1�0�2  �1  �0  � ��� i     #��� I      �/��.�/ .0 changefilepermissions ChangeFilePermissions� ��-� o      �,�, 0 paramstring paramString�-  �.  � k     B�� ��� r     ��� I      �+��*�+ 0 splitstring SplitString� ��� o    �)�) 0 paramstring paramString� ��(� m    �� ���  - , -�(  �*  � J      �� ��� o      �'�'  0 newpermissions newPermissions� ��&� o      �%�% 0 filepath filePath�&  � ��$� Q    B���� k    8�� ��� I   %�#��"
�# .sysoexecTEXT���     TEXT� b    !��� b    ��� m    �� ��� : x a t t r   - d   c o m . a p p l e . q u a r a n t i n e� 1    �!
�! 
spac� n     ��� 1     � 
�  
strq� o    �� 0 filepath filePath�"  � ��� I  & 5���
� .sysoexecTEXT���     TEXT� b   & 1��� b   & -��� b   & +��� b   & )��� m   & '�� ��� 
 c h m o d� 1   ' (�
� 
spac� o   ) *��  0 newpermissions newPermissions� 1   + ,�
� 
spac� n   - 0��� 1   . 0�
� 
strq� o   - .�� 0 filepath filePath�  � ��� L   6 8�� m   6 7�
� boovtrue�  � R      ���
� .ascrerr ****      � ****�  �  � L   @ B�� m   @ A�
� boovfals�$  � ��� l     ����  �  �  � ��� i   $ '��� I      ���� $0 comparemd5hashes CompareMD5Hashes� ��� o      �� 0 paramstring paramString�  �  � k     G�� ��� l     �
���
  � b \ This will check the file integrity of the downloaded template against the known good value.   � ��� �   T h i s   w i l l   c h e c k   t h e   f i l e   i n t e g r i t y   o f   t h e   d o w n l o a d e d   t e m p l a t e   a g a i n s t   t h e   k n o w n   g o o d   v a l u e .� ��� r     ��� I      �	���	 0 splitstring SplitString� ��� o    �� 0 paramstring paramString� ��� m    �� ���  - , -�  �  � J      �� ��� o      �� 0 filepath filePath� ��� o      �� 0 	validhash 	validHash�  � ��� l   ��� �  �  �   � ��� Z    '������� H    �� I    ������� 0 doesfileexist DoesFileExist� ���� o    ���� 0 filepath filePath��  ��  � L   ! #�� m   ! "��
�� boovfals��  ��  � ��� l  ( (��������  ��  ��  � ���� Q   ( G���� k   + =�� ��� r   + 8��� l  + 6������ I  + 6�����
�� .sysoexecTEXT���     TEXT� b   + 2��� b   + .� � m   + , �  m d 5   - q  1   , -��
�� 
spac� n   . 1 1   / 1��
�� 
strq o   . /���� 0 filepath filePath��  ��  ��  � o      ���� 0 checkresult checkResult� �� L   9 = =  9 < o   9 :���� 0 checkresult checkResult o   : ;���� 0 	validhash 	validHash��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   E G		 m   E F��
�� boovfals��  � 

 l     ��������  ��  ��    i   ( + I      ������ 0 copyfile CopyFile �� o      ���� 0 	filepaths 	filePaths��  ��   k     8  l     ����   _ Y Self-explanatory. Copy file from place A to place B. The original file will still exist.    � �   S e l f - e x p l a n a t o r y .   C o p y   f i l e   f r o m   p l a c e   A   t o   p l a c e   B .   T h e   o r i g i n a l   f i l e   w i l l   s t i l l   e x i s t .  r      I      ������ 0 splitstring SplitString  o    ���� 0 	filepaths 	filePaths �� m       �!!  - , -��  ��   J      "" #$# o      ���� 0 
targetfile 
targetFile$ %��% o      ���� "0 destinationfile destinationFile��   &��& Q    8'()' k    .** +,+ I   +��-��
�� .sysoexecTEXT���     TEXT- b    './. b    #010 b    !232 b    454 m    66 �77  c p5 1    ��
�� 
spac3 l    8����8 n     9:9 1     ��
�� 
strq: o    ���� 0 
targetfile 
targetFile��  ��  1 1   ! "��
�� 
spac/ l  # &;����; n   # &<=< 1   $ &��
�� 
strq= o   # $���� "0 destinationfile destinationFile��  ��  ��  , >��> L   , .?? m   , -��
�� boovtrue��  ( R      ������
�� .ascrerr ****      � ****��  ��  ) L   6 8@@ m   6 7��
�� boovfals��   ABA l     ��������  ��  ��  B CDC i   , /EFE I      ��G���� 00 createzipwithlocal7zip CreateZipWithLocal7ZipG H��H o      ���� 0 
zipcommand 
zipCommand��  ��  F Q     IJKI k    LL MNM I   ��O��
�� .sysoexecTEXT���     TEXTO o    ���� 0 
zipcommand 
zipCommand��  N P��P L   	 QQ m   	 
RR �SS  S u c c e s s��  J R      ������
�� .ascrerr ****      � ****��  ��  K L    TT o    ���� 0 errmsg errMsgD UVU l     ��������  ��  ��  V WXW i   0 3YZY I      ��[���� <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiver[ \��\ o      ���� 0 paramstring paramString��  ��  Z k     <]] ^_^ l     ��`a��  ` q k Create a ZIP file of all the PDFs in the target folder. Makes it simpler for you to send them to your KTs.   a �bb �   C r e a t e   a   Z I P   f i l e   o f   a l l   t h e   P D F s   i n   t h e   t a r g e t   f o l d e r .   M a k e s   i t   s i m p l e r   f o r   y o u   t o   s e n d   t h e m   t o   y o u r   K T s ._ cdc r     efe I      ��g���� 0 splitstring SplitStringg hih o    ���� 0 paramstring paramStringi j��j m    kk �ll  - , -��  ��  f J      mm non o      ���� 0 savepath savePatho p��p o      ���� 0 zippath zipPath��  d q��q Q    <rstr k    2uu vwv I   /��x��
�� .sysoexecTEXT���     TEXTx b    +yzy b    ){|{ b    '}~} b    #� b    !��� b    ��� m    �� ���  c d� 1    ��
�� 
spac� n     ��� 1     ��
�� 
strq� o    ���� 0 savepath savePath� m   ! "�� ��� (   & &   / u s r / b i n / z i p   - j  ~ n   # &��� 1   $ &��
�� 
strq� o   # $���� 0 zippath zipPath| 1   ' (��
�� 
spacz m   ) *�� ��� 
 * . p d f��  w ���� L   0 2�� m   0 1�� ���  S u c c e s s��  s R      ������
�� .ascrerr ****      � ****��  ��  t L   : <�� o   : ;���� 0 errmsg errMsg��  X ��� l     ��������  ��  ��  � ��� i   4 7��� I      ������� 0 
deletefile 
DeleteFile� ���� o      ���� 0 filepath filePath��  ��  � k     �� ��� l     ������  � M GSelf-explanatory. This will delete the target file, skipping the Trash.   � ��� � S e l f - e x p l a n a t o r y .   T h i s   w i l l   d e l e t e   t h e   t a r g e t   f i l e ,   s k i p p i n g   t h e   T r a s h .� ��� l      ������  � � � The value of filePath passed to this function is always carefully considered
	(and limited), but at a future point, I will likely add in some safety checks for extra security
	to prevent a dangerous value accidentally being sent to this function.
	   � ����   T h e   v a l u e   o f   f i l e P a t h   p a s s e d   t o   t h i s   f u n c t i o n   i s   a l w a y s   c a r e f u l l y   c o n s i d e r e d 
 	 ( a n d   l i m i t e d ) ,   b u t   a t   a   f u t u r e   p o i n t ,   I   w i l l   l i k e l y   a d d   i n   s o m e   s a f e t y   c h e c k s   f o r   e x t r a   s e c u r i t y 
 	 t o   p r e v e n t   a   d a n g e r o u s   v a l u e   a c c i d e n t a l l y   b e i n g   s e n t   t o   t h i s   f u n c t i o n . 
 	� ���� Q     ���� k    �� ��� I   �����
�� .sysoexecTEXT���     TEXT� b    
��� b    ��� m    �� ��� 
 r m   - f� 1    ��
�� 
spac� l   	������ n    	��� 1    	��
�� 
strq� o    ���� 0 filepath filePath��  ��  ��  � ���� L    �� m    ��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L    �� m    ��
�� boovfals��  � ��� l     ����~��  �  �~  � ��� i   8 ;��� I      �}��|�} "0 doesbundleexist DoesBundleExist� ��{� o      �z�z 0 
bundlepath 
bundlePath�{  �|  � k     �� ��� l     �y���y  � D > Used to check if the Dialog Toolkit Plus script bundle exists   � ��� |   U s e d   t o   c h e c k   i f   t h e   D i a l o g   T o o l k i t   P l u s   s c r i p t   b u n d l e   e x i s t s� ��x� O    ��� L    �� l   ��w�v� I   �u��t
�u .coredoexnull���     ****� 4    �s�
�s 
ditm� o    �r�r 0 
bundlepath 
bundlePath�t  �w  �v  � m     ���                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  �x  � ��� l     �q�p�o�q  �p  �o  � ��� i   < ?��� I      �n��m�n 0 doesfileexist DoesFileExist� ��l� o      �k�k 0 filepath filePath�l  �m  � k     �� ��� l     �j���j  �   Self-explanatory   � ��� "   S e l f - e x p l a n a t o r y� ��i� O    ��� L    �� F    ��� l   ��h�g� I   �f��e
�f .coredoexnull���     ****� 4    �d�
�d 
ditm� o    �c�c 0 filepath filePath�e  �h  �g  � =    ��� n    ��� m    �b
�b 
pcls� 4    �a�
�a 
ditm� o    �`�` 0 filepath filePath� m    �_
�_ 
file� m     ���                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  �i  � ��� l     �^�]�\�^  �]  �\  � ��� i   @ C��� I      �[��Z�[ 0 downloadfile DownloadFile� ��Y� o      �X�X 0 paramstring paramString�Y  �Z  � k     B�� ��� l     �W���W  � Z T Self-explanatory. The value of fileURL is the internet address to the desired file.   � ��� �   S e l f - e x p l a n a t o r y .   T h e   v a l u e   o f   f i l e U R L   i s   t h e   i n t e r n e t   a d d r e s s   t o   t h e   d e s i r e d   f i l e .� ��� r     ��� I      �V �U�V 0 splitstring SplitString   o    �T�T 0 paramstring paramString �S m     �  - , -�S  �U  � J        o      �R�R "0 destinationpath destinationPath 	�Q	 o      �P�P 0 fileurl fileURL�Q  � 
�O
 Q    B k    .  I   +�N�M
�N .sysoexecTEXT���     TEXT b    ' b    # b    ! b     m     �  c u r l   - L   - o 1    �L
�L 
spac l    �K�J n      1     �I
�I 
strq o    �H�H "0 destinationpath destinationPath�K  �J   1   ! "�G
�G 
spac l  # &�F�E n   # & !  1   $ &�D
�D 
strq! o   # $�C�C 0 fileurl fileURL�F  �E  �M   "�B" L   , .## m   , -�A
�A boovtrue�B   R      �@�?�>
�@ .ascrerr ****      � ****�?  �>   k   6 B$$ %&% I  6 ?�='�<
�= .sysodlogaskr        TEXT' b   6 ;()( b   6 9*+* m   6 7,, �-- . E r r o r   d o w n l o a d i n g   f i l e :+ 1   7 8�;
�; 
spac) o   9 :�:�: 0 fileurl fileURL�<  & .�9. L   @ B// m   @ A�8
�8 boovfals�9  �O  � 010 l     �7�6�5�7  �6  �5  1 232 i   D G454 I      �46�3�4 0 findsignature FindSignature6 7�27 o      �1�1 0 signaturepath signaturePath�2  �3  5 k     388 9:9 l     �0;<�0  ; m g If your signature isn't embedded in the Excel file, it will try to find an external JPG or PNG version   < �== �   I f   y o u r   s i g n a t u r e   i s n ' t   e m b e d d e d   i n   t h e   E x c e l   f i l e ,   i t   w i l l   t r y   t o   f i n d   a n   e x t e r n a l   J P G   o r   P N G   v e r s i o n: >�/> Q     3?@A? Z    )BCDEB I    �.F�-�. 0 doesfileexist DoesFileExistF G�,G b    HIH o    �+�+ 0 signaturepath signaturePathI m    JJ �KK  m y S i g n a t u r e . p n g�,  �-  C L    LL b    MNM o    �*�* 0 signaturepath signaturePathN m    OO �PP  m y S i g n a t u r e . p n gD QRQ I    �)S�(�) 0 doesfileexist DoesFileExistS T�'T b    UVU o    �&�& 0 signaturepath signaturePathV m    WW �XX  m y S i g n a t u r e . j p g�'  �(  R Y�%Y L     $ZZ b     #[\[ o     !�$�$ 0 signaturepath signaturePath\ m   ! "]] �^^  m y S i g n a t u r e . p n g�%  E L   ' )__ m   ' (`` �aa  @ R      �#�"�!
�# .ascrerr ****      � ****�"  �!  A L   1 3bb m   1 2cc �dd  �/  3 efe l     � ���   �  �  f ghg i   H Kiji I      �k�� 0 
renamefile 
RenameFilek l�l o      �� 0 paramstring paramString�  �  j k     Dmm non l     �pq�  p z t This pulls double duty for renaming a file or moving it to a new location. (It's the same process to the computer.)   q �rr �   T h i s   p u l l s   d o u b l e   d u t y   f o r   r e n a m i n g   a   f i l e   o r   m o v i n g   i t   t o   a   n e w   l o c a t i o n .   ( I t ' s   t h e   s a m e   p r o c e s s   t o   t h e   c o m p u t e r . )o sts r     uvu I      �w�� 0 splitstring SplitStringw xyx o    �� 0 paramstring paramStringy z�z m    {{ �||  - , -�  �  v J      }} ~~ o      �� 0 
targetfile 
targetFile ��� o      �� 0 newfilename newFilename�  t ��� r    ��� n    ��� 1    �
� 
strq� n    ��� 1    �
� 
psxp� o    �� 0 
targetfile 
targetFile� o      �� 0 
targetfile 
targetFile� ��� r    &��� n    $��� 1   " $�
� 
strq� n    "��� 1     "�
� 
psxp� o     �� 0 newfilename newFilename� o      �
�
 0 newfilename newFilename� ��	� Q   ' D���� k   * :�� ��� I  * 7���
� .sysoexecTEXT���     TEXT� b   * 3��� b   * 1��� b   * /��� b   * -��� m   * +�� ��� 
 m v   - f� 1   + ,�
� 
spac� o   - .�� 0 
targetfile 
targetFile� 1   / 0�
� 
spac� o   1 2�� 0 newfilename newFilename�  � ��� L   8 :�� m   8 9�
� boovtrue�  � R      � ����
�  .ascrerr ****      � ****��  ��  � L   B D�� m   B C��
�� boovfals�	  h ��� l     ��������  ��  ��  � ��� l     ������  �   Folder Manipulation   � ��� (   F o l d e r   M a n i p u l a t i o n� ��� l     ��������  ��  ��  � ��� i   L O��� I      ������� 0 clearfolder ClearFolder� ���� o      ���� 0 foldertoempty folderToEmpty��  ��  � k     ?�� ��� l     ������  � h b Empties the target folder, but only of DOCX, PDF, and ZIP files. This folder will not be deleted.   � ��� �   E m p t i e s   t h e   t a r g e t   f o l d e r ,   b u t   o n l y   o f   D O C X ,   P D F ,   a n d   Z I P   f i l e s .   T h i s   f o l d e r   w i l l   n o t   b e   d e l e t e d .� ���� Q     ?���� k    5�� ��� I   �����
�� .sysoexecTEXT���     TEXT� b    ��� b    ��� b    
��� b    ��� m    �� ���  f i n d� 1    ��
�� 
spac� l   	������ n    	��� 1    	��
�� 
strq� o    ���� 0 foldertoempty folderToEmpty��  ��  � 1   
 ��
�� 
spac� m    �� ��� : - t y p e   f   - n a m e   ' * . p d f '   - d e l e t e��  � ��� I   "�����
�� .sysoexecTEXT���     TEXT� b    ��� b    ��� b    ��� b    ��� m    �� ���  f i n d� 1    ��
�� 
spac� l   ������ n    ��� 1    ��
�� 
strq� o    ���� 0 foldertoempty folderToEmpty��  ��  � 1    ��
�� 
spac� m    �� ��� : - t y p e   f   - n a m e   ' * . z i p '   - d e l e t e��  � ��� I  # 2�����
�� .sysoexecTEXT���     TEXT� b   # .��� b   # ,��� b   # *��� b   # &��� m   # $�� ���  f i n d� 1   $ %��
�� 
spac� l  & )������ n   & )��� 1   ' )��
�� 
strq� o   & '���� 0 foldertoempty folderToEmpty��  ��  � 1   * +��
�� 
spac� m   , -�� ��� < - t y p e   f   - n a m e   ' * . p p t x '   - d e l e t e��  � ���� L   3 5�� m   3 4��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   = ?�� m   = >��
�� boovfals��  � ��� l     ��������  ��  ��  � ��� i   P S��� I      ������� .0 clearpdfsafterzipping ClearPDFsAfterZipping�  ��  o      ���� 0 foldertoempty folderToEmpty��  ��  � Q      k      I   ����
�� .sysoexecTEXT���     TEXT b    	 b    

 b    
 b     m     �  f i n d 1    ��
�� 
spac l   	���� n    	 1    	��
�� 
strq o    ���� 0 foldertoempty folderToEmpty��  ��   1   
 ��
�� 
spac	 m     � : - t y p e   f   - n a m e   ' * . p d f '   - d e l e t e��   �� L     m    ��
�� boovtrue��   R      ������
�� .ascrerr ****      � ****��  ��   L     m    ��
�� boovfals�  l     ��������  ��  ��    i   T W I      �� ���� 0 
copyfolder 
CopyFolder  !��! o      ���� 0 
folderpath 
folderPath��  ��   k     8"" #$# l     ��%&��  % o i Self-explanatory. Copy a folder (or bundle) from place A to place B. The original file will still exist.   & �'' �   S e l f - e x p l a n a t o r y .   C o p y   a   f o l d e r   ( o r   b u n d l e )   f r o m   p l a c e   A   t o   p l a c e   B .   T h e   o r i g i n a l   f i l e   w i l l   s t i l l   e x i s t .$ ()( r     *+* I      ��,���� 0 splitstring SplitString, -.- o    ���� 0 
folderpath 
folderPath. /��/ m    00 �11  - , -��  ��  + J      22 343 o      ���� 0 targetfolder targetFolder4 5��5 o      ���� &0 destinationfolder destinationFolder��  ) 6��6 Q    87897 k    .:: ;<; I   +��=��
�� .sysoexecTEXT���     TEXT= b    '>?> b    #@A@ b    !BCB b    DED m    FF �GG  c p   - R fE 1    ��
�� 
spacC l    H����H n     IJI 1     ��
�� 
strqJ o    ���� 0 targetfolder targetFolder��  ��  A 1   ! "��
�� 
spac? l  # &K����K n   # &LML 1   $ &��
�� 
strqM o   # $���� &0 destinationfolder destinationFolder��  ��  ��  < N��N L   , .OO m   , -��
�� boovtrue��  8 R      ������
�� .ascrerr ****      � ****��  ��  9 L   6 8PP m   6 7��
�� boovfals��   QRQ l     ��������  ��  ��  R STS i   X [UVU I      ��W���� 0 createfolder CreateFolderW X��X o      ���� 0 
folderpath 
folderPath��  ��  V k     YY Z[Z l     ��\]��  \ \ V Self-explanatory. Needed for creating the folder for where the reports will be saved.   ] �^^ �   S e l f - e x p l a n a t o r y .   N e e d e d   f o r   c r e a t i n g   t h e   f o l d e r   f o r   w h e r e   t h e   r e p o r t s   w i l l   b e   s a v e d .[ _��_ Q     `ab` k    cc ded I   ��f��
�� .sysoexecTEXT���     TEXTf b    
ghg b    iji m    kk �ll  m k d i r   - pj 1    ��
�� 
spach l   	m����m n    	non 1    	��
�� 
strqo o    ���� 0 
folderpath 
folderPath��  ��  ��  e p��p L    qq m    ��
�� boovtrue��  a R      ������
�� .ascrerr ****      � ****��  ��  b L    rr m    ��
�� boovfals��  T sts l     ��������  ��  ��  t uvu i   \ _wxw I      ��y���� 0 deletefolder DeleteFoldery z�z o      �~�~ 0 
folderpath 
folderPath�  ��  x k     {{ |}| l     �}~�}  ~ c ] Self-explanatory. Same as with DeleteFile, extra security checks will likely be added later.    ��� �   S e l f - e x p l a n a t o r y .   S a m e   a s   w i t h   D e l e t e F i l e ,   e x t r a   s e c u r i t y   c h e c k s   w i l l   l i k e l y   b e   a d d e d   l a t e r .} ��|� Q     ���� k    �� ��� I   �{��z
�{ .sysoexecTEXT���     TEXT� b    
��� b    ��� m    �� ���  r m   - r f� 1    �y
�y 
spac� l   	��x�w� n    	��� 1    	�v
�v 
strq� o    �u�u 0 
folderpath 
folderPath�x  �w  �z  � ��t� L    �� m    �s
�s boovtrue�t  � R      �r�q�p
�r .ascrerr ****      � ****�q  �p  � L    �� m    �o
�o boovfals�|  v ��� l     �n�m�l�n  �m  �l  � ��� i   ` c��� I      �k��j�k "0 doesfolderexist DoesFolderExist� ��i� o      �h�h 0 
folderpath 
folderPath�i  �j  � k     �� ��� l     �g���g  �   Self-explanatory   � ��� "   S e l f - e x p l a n a t o r y� ��f� O    ��� L    �� F    ��� l   ��e�d� I   �c��b
�c .coredoexnull���     ****� 4    �a�
�a 
ditm� o    �`�` 0 
folderpath 
folderPath�b  �e  �d  � =    ��� n    ��� m    �_
�_ 
pcls� 4    �^�
�^ 
ditm� o    �]�] 0 
folderpath 
folderPath� m    �\
�\ 
cfol� m     ���                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  �f  � ��� l     �[�Z�Y�[  �Z  �Y  � ��� l     �X���X  �   Dialog Boxes   � ���    D i a l o g   B o x e s� ��� l     �W�V�U�W  �V  �U  � ��� i   d g��� I      �T��S�T 80 installdialogdisplayscript InstallDialogDisplayScript� ��R� o      �Q�Q 0 paramstring paramString�R  �S  � k     �� ��� r     ��� b     	��� n     ��� 1    �P
�P 
psxp� l    ��O�N� I    �M��L
�M .earsffdralis        afdr� m     �K
�K afdrcusr�L  �O  �N  � m    �� ��� � L i b r a r y / A p p l i c a t i o n   S c r i p t s / c o m . m i c r o s o f t . E x c e l / D i a l o g D i s p l a y . s c p t� o      �J�J 0 
scriptpath 
scriptPath� ��� r    ��� m    �� ��� � h t t p s : / / r a w . g i t h u b u s e r c o n t e n t . c o m / p a p e r c u t t e r 0 3 2 4 / S p e a k i n g E v a l s / m a i n / D i a l o g D i s p l a y . s c p t� o      �I�I 0 downloadurl downloadURL� ��� l   �H�G�F�H  �G  �F  � ��� l   �E���E  � A ; If an existing version is not found, download a fresh copy   � ��� v   I f   a n   e x i s t i n g   v e r s i o n   i s   n o t   f o u n d ,   d o w n l o a d   a   f r e s h   c o p y� ��� l   �D���D  � e _ Skip this first check until a full update function can be designed. For now, install each time   � ��� �   S k i p   t h i s   f i r s t   c h e c k   u n t i l   a   f u l l   u p d a t e   f u n c t i o n   c a n   b e   d e s i g n e d .   F o r   n o w ,   i n s t a l l   e a c h   t i m e� ��� l   �C���C  � 4 . if DoesFileExist(scriptPath) then return true   � ��� \   i f   D o e s F i l e E x i s t ( s c r i p t P a t h )   t h e n   r e t u r n   t r u e� ��B� L    �� I    �A��@�A 0 downloadfile DownloadFile� ��?� b    ��� b    ��� o    �>�> 0 
scriptpath 
scriptPath� m    �� ���  - , -� o    �=�= 0 downloadurl downloadURL�?  �@  �B  � ��� l     �<�;�:�<  �;  �:  � ��� i   h k��� I      �9��8�9 >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder� ��7� o      �6�6 0 paramstring paramString�7  �8  � k     ]�� ��� r     ��� b     	��� n     ��� 1    �5
�5 
psxp� l     �4�3  I    �2�1
�2 .earsffdralis        afdr m     �0
�0 afdrcusr�1  �4  �3  � m     � 0 L i b r a r y / S c r i p t   L i b r a r i e s� o      �/�/ .0 scriptlibrariesfolder scriptLibrariesFolder�  l   �.�-�,�.  �-  �,   �+ Z    ]�*	 I    �)
�(�) "0 doesfolderexist DoesFolderExist
 �' o    �&�& .0 scriptlibrariesfolder scriptLibrariesFolder�'  �(   L     o    �%�% .0 scriptlibrariesfolder scriptLibrariesFolder�*  	 Q    ] k    Q  l   �$�$   m g ~/Library is typically a read-only folder, so I need to requst your password to create the need folder    � �   ~ / L i b r a r y   i s   t y p i c a l l y   a   r e a d - o n l y   f o l d e r ,   s o   I   n e e d   t o   r e q u s t   y o u r   p a s s w o r d   t o   c r e a t e   t h e   n e e d   f o l d e r  I   *�#
�# .sysoexecTEXT���     TEXT b    $ b      m     �  m k d i r   - p 1    �"
�" 
spac n     # !  1   ! #�!
�! 
strq! o     !� �  .0 scriptlibrariesfolder scriptLibrariesFolder �"�
� 
badm" m   % &�
� boovtrue�   #$# l  + +�%&�  % %  Set your username as the owner   & �'' >   S e t   y o u r   u s e r n a m e   a s   t h e   o w n e r$ ()( I  + B�*+
� .sysoexecTEXT���     TEXT* b   + <,-, b   + 8./. b   + 6010 m   + ,22 �33  c h o w n  1 n   , 5454 1   3 5�
� 
strq5 l  , 36��6 n   , 3787 1   1 3�
� 
sisn8 l  , 19��9 I  , 1���
� .sysosigtsirr   ��� null�  �  �  �  �  �  / 1   6 7�
� 
spac- n   8 ;:;: 1   9 ;�
� 
strq; o   8 9�� .0 scriptlibrariesfolder scriptLibrariesFolder+ �<�
� 
badm< m   = >�
� boovtrue�  ) =>= l  C C�?@�  ? 5 / Give your username READ and WRITE permissions.   @ �AA ^   G i v e   y o u r   u s e r n a m e   R E A D   a n d   W R I T E   p e r m i s s i o n s .> BCB I  C N�
DE
�
 .sysoexecTEXT���     TEXTD b   C HFGF m   C DHH �II  c h m o d   u + r w  G n   D GJKJ 1   E G�	
�	 
strqK o   D E�� .0 scriptlibrariesfolder scriptLibrariesFolderE �L�
� 
badmL m   I J�
� boovtrue�  C M�M L   O QNN o   O P�� .0 scriptlibrariesfolder scriptLibrariesFolder�   R      ��� 
� .ascrerr ****      � ****�  �    L   Y ]OO m   Y \PP �QQ  �+  � RSR l     ��������  ��  ��  S TUT i   l oVWV I      ��X���� 40 installdialogtoolkitplus InstallDialogToolkitPlusX Y��Y o      ���� "0 resourcesfolder resourcesFolder��  ��  W k     �ZZ [\[ r     ]^] m     __ �`` � h t t p s : / / r a w . g i t h u b u s e r c o n t e n t . c o m / p a p e r c u t t e r 0 3 2 4 / S p e a k i n g E v a l s / m a i n / D i a l o g _ T o o l k i t . z i p^ o      ���� 0 downloadurl downloadURL\ aba r    cdc b    efe n    ghg 1   	 ��
�� 
psxph l   	i����i I   	��j��
�� .earsffdralis        afdrj m    ��
�� afdrcusr��  ��  ��  f m    kk �ll 0 L i b r a r y / S c r i p t   L i b r a r i e sd o      ���� .0 scriptlibrariesfolder scriptLibrariesFolderb mnm r    opo m    qq �rr 4 / D i a l o g   T o o l k i t   P l u s . s c p t dp o      ���� $0 dialogbundlename dialogBundleNamen sts r    uvu b    wxw o    ���� .0 scriptlibrariesfolder scriptLibrariesFolderx o    ���� $0 dialogbundlename dialogBundleNamev o      ���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundlet yzy r    {|{ b    }~} o    ���� "0 resourcesfolder resourcesFolder~ m     ��� & / D i a l o g _ T o o l k i t . z i p| o      ���� 0 zipfilepath zipFilePathz ��� r     %��� b     #��� o     !���� "0 resourcesfolder resourcesFolder� m   ! "�� ��� $ / d i a l o g T o o l k i t T e m p� o      ���� &0 zipextractionpath zipExtractionPath� ��� l  & &��������  ��  ��  � ��� l  & &������  � 0 * Initial check to see if already installed   � ��� T   I n i t i a l   c h e c k   t o   s e e   i f   a l r e a d y   i n s t a l l e d� ��� Z  & 5������� I   & ,������� "0 doesbundleexist DoesBundleExist� ���� o   ' (���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  � L   / 1�� m   / 0��
�� boovtrue��  ��  � ��� l  6 6��������  ��  ��  � ��� l  6 6������  � 3 - Ensure resources folder exists for later use   � ��� Z   E n s u r e   r e s o u r c e s   f o l d e r   e x i s t s   f o r   l a t e r   u s e� ��� Z   6 W������� H   6 =�� I   6 <������� "0 doesfolderexist DoesFolderExist� ���� o   7 8���� "0 resourcesfolder resourcesFolder��  ��  � Q   @ S���� I   C I������� 0 createfolder CreateFolder� ���� o   D E���� "0 resourcesfolder resourcesFolder��  ��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   Q S�� m   Q R��
�� boovfals��  ��  � ��� l  X X��������  ��  ��  � ��� l  X X������  � G A Check for a local copy and move it to the needed folder if found   � ��� �   C h e c k   f o r   a   l o c a l   c o p y   a n d   m o v e   i t   t o   t h e   n e e d e d   f o l d e r   i f   f o u n d� ��� Z   X |������� I   X `������� "0 doesbundleexist DoesBundleExist� ���� b   Y \��� o   Y Z���� "0 resourcesfolder resourcesFolder� o   Z [���� $0 dialogbundlename dialogBundleName��  ��  � Z   c x������� I   c o������� 0 
copyfolder 
CopyFolder� ���� b   d k��� b   d i��� b   d g��� o   d e���� "0 resourcesfolder resourcesFolder� o   e f���� $0 dialogbundlename dialogBundleName� m   g h�� ���  - , -� o   i j���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  � L   r t�� m   r s��
�� boovtrue��  ��  ��  ��  � ��� l  } }��������  ��  ��  � ��� l  } }������  � !  Otherwise, download and...   � ��� 6   O t h e r w i s e ,   d o w n l o a d   a n d . . .� ��� Z   } �������� I   } �������� 0 downloadfile DownloadFile� ���� b   ~ ���� b   ~ ���� o   ~ ���� 0 zipfilepath zipFilePath� m    ��� ���  - , -� o   � ����� 0 downloadurl downloadURL��  ��  � Q   � ������ k   � ��� ��� l  � �������  �   ...extract the files...   � ��� 0   . . . e x t r a c t   t h e   f i l e s . . .� ��� I  � ������
�� .sysoexecTEXT���     TEXT� b   � ���� b   � ���� b   � ���� b   � ���� m   � ��� ���  u n z i p   - o� 1   � ���
�� 
spac� l  � ������� n   � ���� 1   � ���
�� 
strq� o   � ����� 0 zipfilepath zipFilePath��  ��  � m   � ��� ���    - d  � l  � ������� n   � ���� 1   � ���
�� 
strq� o   � ����� &0 zipextractionpath zipExtractionPath��  ��  ��  � ��� l  � �������  � 6 0 ...keep a local copy in the resources folder...   � ��� `   . . . k e e p   a   l o c a l   c o p y   i n   t h e   r e s o u r c e s   f o l d e r . . .� ��� I   � �������� 0 
copyfolder 
CopyFolder� ���� b   � �   b   � � b   � � b   � � b   � �	 o   � ����� &0 zipextractionpath zipExtractionPath	 m   � �

 �  / D i a l o g _ T o o l k i t o   � ����� $0 dialogbundlename dialogBundleName m   � � �  - , - o   � ����� "0 resourcesfolder resourcesFolder o   � ����� $0 dialogbundlename dialogBundleName��  ��  �  l  � �����   ; 5 ...and copy the script bundle to the required folder    � j   . . . a n d   c o p y   t h e   s c r i p t   b u n d l e   t o   t h e   r e q u i r e d   f o l d e r �� I   � ������� 0 
copyfolder 
CopyFolder �� b   � � b   � � b   � � b   � � o   � ����� &0 zipextractionpath zipExtractionPath m   � � �  / D i a l o g _ T o o l k i t o   � ����� $0 dialogbundlename dialogBundleName m   � �   �!!  - , - o   � ����� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  ��  � R      ������
�� .ascrerr ****      � ****��  ��  ��  ��  ��  � "#" l  � ���������  ��  ��  # $%$ l  � ���&'��  & D > Remove unneeded files and folders created during this process   ' �(( |   R e m o v e   u n n e e d e d   f i l e s   a n d   f o l d e r s   c r e a t e d   d u r i n g   t h i s   p r o c e s s% )*) I   � ���+���� 0 
deletefile 
DeleteFile+ ,��, o   � ����� 0 zipfilepath zipFilePath��  ��  * -.- I   � ���/���� 0 deletefolder DeleteFolder/ 0��0 o   � ����� &0 zipextractionpath zipExtractionPath��  ��  . 121 l  � �����~��  �  �~  2 343 l  � ��}56�}  5 V P One final check to verify installation was successful and return true if it was   6 �77 �   O n e   f i n a l   c h e c k   t o   v e r i f y   i n s t a l l a t i o n   w a s   s u c c e s s f u l   a n d   r e t u r n   t r u e   i f   i t   w a s4 8�|8 L   � �99 I   � ��{:�z�{ "0 doesbundleexist DoesBundleExist: ;�y; o   � ��x�x 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�y  �z  �|  U <=< l     �w�v�u�w  �v  �u  = >?> i   p s@A@ I      �tB�s�t 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlusB C�rC o      �q�q "0 resourcesfolder resourcesFolder�r  �s  A k     UDD EFE r     GHG b     	IJI n     KLK 1    �p
�p 
psxpL l    M�o�nM I    �mN�l
�m .earsffdralis        afdrN m     �k
�k afdrcusr�l  �o  �n  J m    OO �PP d L i b r a r y / S c r i p t   L i b r a r i e s / D i a l o g   T o o l k i t   P l u s . s c p t dH o      �j�j 20 dialogtoolkitplusbundle dialogToolkitPlusBundleF QRQ r    STS b    UVU o    �i�i "0 resourcesfolder resourcesFolderV m    WW �XX 4 / D i a l o g   T o o l k i t   P l u s . s c p t dT o      �h�h 0 	localcopy 	localCopyR YZY l   �g�f�e�g  �f  �e  Z [\[ Z    R]^�d_] I    �c`�b�c "0 doesbundleexist DoesBundleExist` a�aa o    �`�` 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�a  �b  ^ Q    Lbcdb k    Aee fgf Z   6hi�_�^h H    %jj I    $�]k�\�] "0 doesbundleexist DoesBundleExistk l�[l o     �Z�Z 0 	localcopy 	localCopy�[  �\  i I   ( 2�Ym�X�Y 0 
copyfolder 
CopyFolderm n�Wn b   ) .opo b   ) ,qrq o   ) *�V�V 20 dialogtoolkitplusbundle dialogToolkitPlusBundler m   * +ss �tt  - , -p o   , -�U�U 0 	localcopy 	localCopy�W  �X  �_  �^  g uvu I   7 =�Tw�S�T 0 deletefolder DeleteFolderw x�Rx o   8 9�Q�Q 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�R  �S  v y�Py r   > Az{z m   > ?�O
�O boovtrue{ o      �N�N 0 removalresult removalResult�P  c R      �M�L�K
�M .ascrerr ****      � ****�L  �K  d r   I L|}| m   I J�J
�J boovfals} o      �I�I 0 removalresult removalResult�d  _ r   O R~~ m   O P�H
�H boovtrue o      �G�G 0 removalresult removalResult\ ��� l  S S�F�E�D�F  �E  �D  � ��C� L   S U�� o   S T�B�B 0 removalresult removalResult�C  ? ��A� l     �@�?�>�@  �?  �>  �A       �=�������������������������������=  � �<�;�:�9�8�7�6�5�4�3�2�1�0�/�.�-�,�+�*�)�(�'�&�%�$�#�"�!� �< 00 getscriptversionnumber GetScriptVersionNumber�; "0 getmacosversion GetMacOSVersion�: 80 checkaccessibilitysettings CheckAccessibilitySettings�9 0 splitstring SplitString�8 "0 loadapplication LoadApplication�7 0 isapploaded IsAppLoaded�6 "0 closepowerpoint ClosePowerPoint�5 0 	closeword 	CloseWord�4 .0 changefilepermissions ChangeFilePermissions�3 $0 comparemd5hashes CompareMD5Hashes�2 0 copyfile CopyFile�1 00 createzipwithlocal7zip CreateZipWithLocal7Zip�0 <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiver�/ 0 
deletefile 
DeleteFile�. "0 doesbundleexist DoesBundleExist�- 0 doesfileexist DoesFileExist�, 0 downloadfile DownloadFile�+ 0 findsignature FindSignature�* 0 
renamefile 
RenameFile�) 0 clearfolder ClearFolder�( .0 clearpdfsafterzipping ClearPDFsAfterZipping�' 0 
copyfolder 
CopyFolder�& 0 createfolder CreateFolder�% 0 deletefolder DeleteFolder�$ "0 doesfolderexist DoesFolderExist�# 80 installdialogdisplayscript InstallDialogDisplayScript�" >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder�! 40 installdialogtoolkitplus InstallDialogToolkitPlus�  80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus� � ������ 00 getscriptversionnumber GetScriptVersionNumber� ��� �  �� 0 paramstring paramString�  � �� 0 paramstring paramString� �� 4��� �� � &������ "0 getmacosversion GetMacOSVersion� ��� �  �� 0 paramstring paramString�  � ��� 0 paramstring paramString� 0 	osversion 	osVersion�  8���
� .sysoexecTEXT���     TEXT�  �  �  �j E�O�W X  h� � A��
���	� 80 checkaccessibilitysettings CheckAccessibilitySettings� ��� �  �� 0 
apptocheck 
appToCheck�
  � ��� 0 
apptocheck 
appToCheck� ,0 accessibilityenabled accessibilityEnabled�  n������ ��������
� 
prcs
� 
pnam�  
� 
pvis
� 
pcap
�  
uiel
�� 
enaB
�� 
bool��  ��  �	 6 -� %*�-�,�[�,\Ze81�	 *�/�-�,E�&E�O�UW 	X 	 
f� �� |���������� 0 splitstring SplitString�� ����� �  ������ &0 passedparamstring passedParamString�� (0 parameterseparator parameterSeparator��  � ���������� &0 passedparamstring passedParamString�� (0 parameterseparator parameterSeparator�� 00 oldtextitemsdelimiters oldTextItemsDelimiters�� *0 separatedparameters separatedParameters� ������
�� 
ascr
�� 
txdl
�� 
citm��  � *�,E�O�*�,FO��-E�O�*�,FUO�� �� ����������� "0 loadapplication LoadApplication�� ����� �  ���� 0 appname appName��  � �������� 0 appname appName�� 0 errmsg errMsg�� 0 errnum errNum� 	���� ���� ��� � �
�� 
capp
�� .miscactvnull��� ��� null�� 0 errmsg errMsg� ������
�� 
errn�� 0 errnum errNum��  
�� 
spac�� * *�/ *j UO�W X  ��%�%�%�%�%�%� �� ����������� 0 isapploaded IsAppLoaded�� ����� �  ���� 0 appname appName��  � ���������� 0 appname appName�� 0 
loadresult 
loadResult�� 0 errmsg errMsg�� 0 errnum errNum� ������ ����
�� 
prcs
�� 
pnam
�� 
spac�� 0 errmsg errMsg� ������
�� 
errn�� 0 errnum errNum��  �� ; (� *�-�,� ��%�%E�Y 	��%�%E�UO�W X  �%�%�%�%�%� ��#���������� "0 closepowerpoint ClosePowerPoint�� ����� �  ���� 0 paramstring paramString��  � ������ 0 paramstring paramString�� 0 closeresult closeResult� P����=D��HL����R
�� 
prcs
�� 
pnam
�� .aevtquitnull��� ��� null��  ��  �� 4 +� #*�-�,� � *j UO�E�Y �E�O�UW 	X  	�� ��Y���������� 0 	closeword 	CloseWord�� ����� �  ���� 0 paramstring paramString��  � ������ 0 paramstring paramString�� 0 closeresult closeResult� �����sz��~������
�� 
prcs
�� 
pnam
�� .aevtquitnull��� ��� null��  ��  �� 4 +� #*�-�,� � *j UO�E�Y �E�O�UW 	X  	�� ������������� .0 changefilepermissions ChangeFilePermissions�� ����� �  ���� 0 paramstring paramString��  � �������� 0 paramstring paramString��  0 newpermissions newPermissions�� 0 filepath filePath� 
������������������� 0 splitstring SplitString
�� 
cobj
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� C*��l+ E[�k/E�Z[�l/E�ZO #��%��,%j O��%�%�%��,%j OeW 	X  	f� ������������� $0 comparemd5hashes CompareMD5Hashes�� ����� �  ���� 0 paramstring paramString��  � ���������� 0 paramstring paramString�� 0 filepath filePath�� 0 	validhash 	validHash�� 0 checkresult checkResult� 
������������������� 0 splitstring SplitString
�� 
cobj�� 0 doesfileexist DoesFileExist
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� H*��l+ E[�k/E�Z[�l/E�ZO*�k+  fY hO ��%��,%j E�O�� W 	X  	f� ������������ 0 copyfile CopyFile�� ����� �  ���� 0 	filepaths 	filePaths��  � �������� 0 	filepaths 	filePaths�� 0 
targetfile 
targetFile�� "0 destinationfile destinationFile� 	 ����6������������ 0 splitstring SplitString
�� 
cobj
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� 9*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW 	X  f� ��F��~���}�� 00 createzipwithlocal7zip CreateZipWithLocal7Zip� �|��| �  �{�{ 0 
zipcommand 
zipCommand�~  � �z�y�z 0 
zipcommand 
zipCommand�y 0 errmsg errMsg� �xR�w�v
�x .sysoexecTEXT���     TEXT�w  �v  �}  �j  O�W 	X  �� �uZ�t�s���r�u <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiver�t �q��q �  �p�p 0 paramstring paramString�s  � �o�n�m�l�o 0 paramstring paramString�n 0 savepath savePath�m 0 zippath zipPath�l 0 errmsg errMsg� k�k�j��i�h���g��f�e�k 0 splitstring SplitString
�j 
cobj
�i 
spac
�h 
strq
�g .sysoexecTEXT���     TEXT�f  �e  �r =*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%�%�%j O�W 	X 
 �� �d��c�b���a�d 0 
deletefile 
DeleteFile�c �`��` �  �_�_ 0 filepath filePath�b  � �^�^ 0 filepath filePath� ��]�\�[�Z�Y
�] 
spac
�\ 
strq
�[ .sysoexecTEXT���     TEXT�Z  �Y  �a  ��%��,%j OeW 	X  f� �X��W�V���U�X "0 doesbundleexist DoesBundleExist�W �T��T �  �S�S 0 
bundlepath 
bundlePath�V  � �R�R 0 
bundlepath 
bundlePath� ��Q�P
�Q 
ditm
�P .coredoexnull���     ****�U � *�/j U� �O��N�M���L�O 0 doesfileexist DoesFileExist�N �K��K �  �J�J 0 filepath filePath�M  � �I�I 0 filepath filePath� ��H�G�F�E�D
�H 
ditm
�G .coredoexnull���     ****
�F 
pcls
�E 
file
�D 
bool�L � *�/j 	 *�/�,� �&U� �C��B�A���@�C 0 downloadfile DownloadFile�B �?��? �  �>�> 0 paramstring paramString�A  � �=�<�;�= 0 paramstring paramString�< "0 destinationpath destinationPath�; 0 fileurl fileURL� �:�9�8�7�6�5�4,�3�: 0 splitstring SplitString
�9 
cobj
�8 
spac
�7 
strq
�6 .sysoexecTEXT���     TEXT�5  �4  
�3 .sysodlogaskr        TEXT�@ C*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW X  ��%�%j 
Of� �25�1�0���/�2 0 findsignature FindSignature�1 �.��. �  �-�- 0 signaturepath signaturePath�0  � �,�, 0 signaturepath signaturePath� 	J�+OW]`�*�)c�+ 0 doesfileexist DoesFileExist�*  �)  �/ 4 +*��%k+  	��%Y *��%k+  	��%Y �W 	X  �� �(j�'�&���%�( 0 
renamefile 
RenameFile�' �$��$ �  �#�# 0 paramstring paramString�&  � �"�!� �" 0 paramstring paramString�! 0 
targetfile 
targetFile�  0 newfilename newFilename� 
{���������� 0 splitstring SplitString
� 
cobj
� 
psxp
� 
strq
� 
spac
� .sysoexecTEXT���     TEXT�  �  �% E*��l+ E[�k/E�Z[�l/E�ZO��,�,E�O��,�,E�O ��%�%�%�%j OeW 	X  	f� �������� 0 clearfolder ClearFolder� ��� �  �� 0 foldertoempty folderToEmpty�  � �� 0 foldertoempty folderToEmpty� �����������
� 
spac
� 
strq
� .sysoexecTEXT���     TEXT�  �  � @ 7��%��,%�%�%j O��%��,%�%�%j O��%��,%�%�%j OeW 	X 	 
f� ���
�	���� .0 clearpdfsafterzipping ClearPDFsAfterZipping�
 ��� �  �� 0 foldertoempty folderToEmpty�	  � �� 0 foldertoempty folderToEmpty� ����� 
� 
spac
� 
strq
� .sysoexecTEXT���     TEXT�  �   �   ��%��,%�%�%j OeW 	X  f� ������������ 0 
copyfolder 
CopyFolder�� ����� �  ���� 0 
folderpath 
folderPath��  � �������� 0 
folderpath 
folderPath�� 0 targetfolder targetFolder�� &0 destinationfolder destinationFolder� 	0����F������������ 0 splitstring SplitString
�� 
cobj
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� 9*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW 	X  f� ��V���������� 0 createfolder CreateFolder�� ����� �  ���� 0 
folderpath 
folderPath��  � ���� 0 
folderpath 
folderPath� k����������
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  ��  ��%��,%j OeW 	X  f� ��x���������� 0 deletefolder DeleteFolder�� ����� �  ���� 0 
folderpath 
folderPath��  � ���� 0 
folderpath 
folderPath� �����������
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  ��  ��%��,%j OeW 	X  f� ������������� "0 doesfolderexist DoesFolderExist�� ����� �  ���� 0 
folderpath 
folderPath��  � ���� 0 
folderpath 
folderPath� �����������
�� 
ditm
�� .coredoexnull���     ****
�� 
pcls
�� 
cfol
�� 
bool�� � *�/j 	 *�/�,� �&U� ������������� 80 installdialogdisplayscript InstallDialogDisplayScript�� ����� �  ���� 0 paramstring paramString��  � �������� 0 paramstring paramString�� 0 
scriptpath 
scriptPath�� 0 downloadurl downloadURL� �����������
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp�� 0 downloadfile DownloadFile�� �j �,�%E�O�E�O*��%�%k+ � ������������� >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder�� ����� �  ���� 0 paramstring paramString��  � ������ 0 paramstring paramString�� .0 scriptlibrariesfolder scriptLibrariesFolder� ����������������2����H����P
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp�� "0 doesfolderexist DoesFolderExist
�� 
spac
�� 
strq
�� 
badm
�� .sysoexecTEXT���     TEXT
�� .sysosigtsirr   ��� null
�� 
sisn��  ��  �� ^�j �,�%E�O*�k+  �Y E 9��%��,%�el 	O�*j �,�,%�%��,%�el 	O���,%�el 	O�W X  a � ��W���������� 40 installdialogtoolkitplus InstallDialogToolkitPlus�� ����� �  ���� "0 resourcesfolder resourcesFolder��  � ���������������� "0 resourcesfolder resourcesFolder�� 0 downloadurl downloadURL�� .0 scriptlibrariesfolder scriptLibrariesFolder�� $0 dialogbundlename dialogBundleName�� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�� 0 zipfilepath zipFilePath�� &0 zipextractionpath zipExtractionPath� _������kq�������������������������
 ����
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp�� "0 doesbundleexist DoesBundleExist�� "0 doesfolderexist DoesFolderExist�� 0 createfolder CreateFolder��  ��  �� 0 
copyfolder 
CopyFolder�� 0 downloadfile DownloadFile
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT�� 0 
deletefile 
DeleteFile�� 0 deletefolder DeleteFolder�� ��E�O�j �,�%E�O�E�O��%E�O��%E�O��%E�O*�k+  eY hO*�k+ 	  *�k+ 
W 	X  fY hO*��%k+  *��%�%�%k+  eY hY hO*��%�%k+  T Ha _ %�a ,%a %�a ,%j O*�a %�%a %�%�%k+ O*�a %�%a %�%k+ W X  hY hO*�k+ O*�k+ O*�k+ � ��A���������� 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus�� ����� �  ���� "0 resourcesfolder resourcesFolder��  � ���������� "0 resourcesfolder resourcesFolder�� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�� 0 	localcopy 	localCopy�� 0 removalresult removalResult� ������OW��s����~�}
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp�� "0 doesbundleexist DoesBundleExist�� 0 
copyfolder 
CopyFolder� 0 deletefolder DeleteFolder�~  �}  �� V�j �,�%E�O��%E�O*�k+  6 (*�k+  *��%�%k+ Y hO*�k+ OeE�W 
X 	 
fE�Y eE�O�ascr  ��ޭ