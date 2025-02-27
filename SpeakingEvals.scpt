FasdUAS 1.101.10   ��   ��    k             l      ��  ��    � |
Helper Scripts for the DYB Speaking Evaluations Excel spreadsheet

Version: 1.3.1
Build:   20250227
Warren Feltmate
� 2025
     � 	 	 � 
 H e l p e r   S c r i p t s   f o r   t h e   D Y B   S p e a k i n g   E v a l u a t i o n s   E x c e l   s p r e a d s h e e t 
 
 V e r s i o n :   1 . 3 . 1 
 B u i l d :       2 0 2 5 0 2 2 7 
 W a r r e n   F e l t m a t e 
 �   2 0 2 5 
   
  
 l     ��������  ��  ��        l     ��  ��      Environment Variables     �   ,   E n v i r o n m e n t   V a r i a b l e s      l     ��������  ��  ��        i         I      �� ���� 00 getscriptversionnumber GetScriptVersionNumber   ��  o      ���� 0 paramstring paramString��  ��    k            l     ��  ��    ? 9- Use build number to determine if an update is available     �   r -   U s e   b u i l d   n u m b e r   t o   d e t e r m i n e   i f   a n   u p d a t e   i s   a v a i l a b l e   ��  L          m     ���� 4�s��     ! " ! l     ��������  ��  ��   "  # $ # i     % & % I      �� '���� "0 getmacosversion GetMacOSVersion '  (�� ( o      ���� 0 paramstring paramString��  ��   & k      ) )  * + * l     �� , -��   , ` Z Not currently used, but could be helpful if there are issues with older versions of MacOS    - � . . �   N o t   c u r r e n t l y   u s e d ,   b u t   c o u l d   b e   h e l p f u l   i f   t h e r e   a r e   i s s u e s   w i t h   o l d e r   v e r s i o n s   o f   M a c O S +  /�� / Q      0 1�� 0 k     2 2  3 4 3 r    
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
appToCheck��  ��   Y o      ���� ,0 accessibilityenabled accessibilityEnabled W  l�� l L   ( * m m o   ( )���� ,0 accessibilityenabled accessibilityEnabled��   O m     n n�                                                                                  sevs  alis    N  macOS                      �z2[BD ����System Events.app                                              �����z2[        ����  
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
loadResult � m    �                                                                                  sevs  alis    N  macOS                      �z2[BD ����System Events.app                                              �����z2[        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��   � �r L   $ &		 o   $ %�q�q 0 
loadresult 
loadResult�r   � R      �p

�p .ascrerr ****      � ****
 o      �o�o 0 errmsg errMsg �n�m
�n 
errn o      �l�l 0 errnum errNum�m   � L   . : b   . 9 b   . 7 b   . 5 b   . 3 b   . 1 m   . / �  E r r o r   l o a d i n g   o   / 0�k�k 0 appname appName m   1 2 �  :   o   3 4�j�j 0 errnum errNum m   5 6 �    -   o   7 8�i�i 0 errmsg errMsg�   �  l     �h�g�f�h  �g  �f    !  i    "#" I      �e$�d�e 0 	closeword 	CloseWord$ %�c% o      �b�b 0 paramstring paramString�c  �d  # k     3&& '(' l     �a)*�a  ) u o This will completely close MS Word, even from the Dock. This reduces the chances of errors on subsequent runs.   * �++ �   T h i s   w i l l   c o m p l e t e l y   c l o s e   M S   W o r d ,   e v e n   f r o m   t h e   D o c k .   T h i s   r e d u c e s   t h e   c h a n c e s   o f   e r r o r s   o n   s u b s e q u e n t   r u n s .( ,�`, Q     3-./- O    )010 k    (22 343 Z    %56�_75 E    898 l   :�^�]: n    ;<; 1   
 �\
�\ 
pnam< 2    
�[
�[ 
prcs�^  �]  9 m    == �>>  M i c r o s o f t   W o r d6 k    ?? @A@ O   BCB I   �Z�Y�X
�Z .aevtquitnull��� ��� null�Y  �X  C m    DD�                                                                                  MSWD  alis    4  macOS                      �z2[BD ����Microsoft Word.app                                             ������5        ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    m a c O S  Applications/Microsoft Word.app   / ��  A E�WE r    FGF m    HH �II D W o r d   h a s   s u c c e s s f u l l y   b e e n   c l o s e d .G o      �V�V 0 closeresult closeResult�W  �_  7 r   " %JKJ m   " #LL �MM < W o r d   i s   n o t   c u r r e n t l y   r u n n i n g .K o      �U�U 0 closeresult closeResult4 N�TN L   & (OO o   & '�S�S 0 closeresult closeResult�T  1 m    PP�                                                                                  sevs  alis    N  macOS                      �z2[BD ����System Events.app                                              �����z2[        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  . R      �R�Q�P
�R .ascrerr ****      � ****�Q  �P  / L   1 3QQ m   1 2RR �SS P T h e r e   w a s   a n   e r r o r   t r y i n g   t o   c l o s e   W o r d .�`  ! TUT l     �O�N�M�O  �N  �M  U VWV l     �LXY�L  X   File Manipulation   Y �ZZ $   F i l e   M a n i p u l a t i o nW [\[ l     �K�J�I�K  �J  �I  \ ]^] i    _`_ I      �Ha�G�H .0 changefilepermissions ChangeFilePermissionsa b�Fb o      �E�E 0 paramstring paramString�F  �G  ` k     Bcc ded r     fgf I      �Dh�C�D 0 splitstring SplitStringh iji o    �B�B 0 paramstring paramStringj k�Ak m    ll �mm  - , -�A  �C  g J      nn opo o      �@�@  0 newpermissions newPermissionsp q�?q o      �>�> 0 filepath filePath�?  e r�=r Q    Bstus k    8vv wxw I   %�<y�;
�< .sysoexecTEXT���     TEXTy b    !z{z b    |}| m    ~~ � : x a t t r   - d   c o m . a p p l e . q u a r a n t i n e} 1    �:
�: 
spac{ n     ��� 1     �9
�9 
strq� o    �8�8 0 filepath filePath�;  x ��� I  & 5�7��6
�7 .sysoexecTEXT���     TEXT� b   & 1��� b   & -��� b   & +��� b   & )��� m   & '�� ��� 
 c h m o d� 1   ' (�5
�5 
spac� o   ) *�4�4  0 newpermissions newPermissions� 1   + ,�3
�3 
spac� n   - 0��� 1   . 0�2
�2 
strq� o   - .�1�1 0 filepath filePath�6  � ��0� L   6 8�� m   6 7�/
�/ boovtrue�0  t R      �.�-�,
�. .ascrerr ****      � ****�-  �,  u L   @ B�� m   @ A�+
�+ boovfals�=  ^ ��� l     �*�)�(�*  �)  �(  � ��� i     #��� I      �'��&�' $0 comparemd5hashes CompareMD5Hashes� ��%� o      �$�$ 0 paramstring paramString�%  �&  � k     G�� ��� l     �#���#  � b \ This will check the file integrity of the downloaded template against the known good value.   � ��� �   T h i s   w i l l   c h e c k   t h e   f i l e   i n t e g r i t y   o f   t h e   d o w n l o a d e d   t e m p l a t e   a g a i n s t   t h e   k n o w n   g o o d   v a l u e .� ��� r     ��� I      �"��!�" 0 splitstring SplitString� ��� o    � �  0 paramstring paramString� ��� m    �� ���  - , -�  �!  � J      �� ��� o      �� 0 filepath filePath� ��� o      �� 0 	validhash 	validHash�  � ��� l   ����  �  �  � ��� Z    '����� H    �� I    ���� 0 doesfileexist DoesFileExist� ��� o    �� 0 filepath filePath�  �  � L   ! #�� m   ! "�
� boovfals�  �  � ��� l  ( (����  �  �  � ��� Q   ( G���� k   + =�� ��� r   + 8��� l  + 6���� I  + 6���

� .sysoexecTEXT���     TEXT� b   + 2��� b   + .��� m   + ,�� ���  m d 5   - q� 1   , -�	
�	 
spac� n   . 1��� 1   / 1�
� 
strq� o   . /�� 0 filepath filePath�
  �  �  � o      �� 0 checkresult checkResult� ��� L   9 =�� =  9 <��� o   9 :�� 0 checkresult checkResult� o   : ;�� 0 	validhash 	validHash�  � R      ��� 
� .ascrerr ****      � ****�  �   � L   E G�� m   E F��
�� boovfals�  � ��� l     ��������  ��  ��  � ��� i   $ '��� I      ������� 0 copyfile CopyFile� ���� o      ���� 0 	filepaths 	filePaths��  ��  � k     8�� ��� l     ������  � _ Y Self-explanatory. Copy file from place A to place B. The original file will still exist.   � ��� �   S e l f - e x p l a n a t o r y .   C o p y   f i l e   f r o m   p l a c e   A   t o   p l a c e   B .   T h e   o r i g i n a l   f i l e   w i l l   s t i l l   e x i s t .� ��� r     ��� I      ������� 0 splitstring SplitString� ��� o    ���� 0 	filepaths 	filePaths� ���� m    �� ���  - , -��  ��  � J      �� ��� o      ���� 0 
targetfile 
targetFile� ���� o      ���� "0 destinationfile destinationFile��  � ���� Q    8���� k    .�� ��� I   +�����
�� .sysoexecTEXT���     TEXT� b    '��� b    #��� b    !��� b    ��� m       �  c p� 1    ��
�� 
spac� l    ���� n      1     ��
�� 
strq o    ���� 0 
targetfile 
targetFile��  ��  � 1   ! "��
�� 
spac� l  # &���� n   # & 1   $ &��
�� 
strq o   # $���� "0 destinationfile destinationFile��  ��  ��  � �� L   , .		 m   , -��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   6 8

 m   6 7��
�� boovfals��  �  l     ��������  ��  ��    i   ( + I      ������ (0 createzipwithditto CreateZipWithDitto �� o      ���� 0 
zipcommand 
zipCommand��  ��   Q      k      I   ����
�� .sysoexecTEXT���     TEXT o    ���� 0 
zipcommand 
zipCommand��   �� L   	  m   	 
 �  S u c c e s s��   R      ������
�� .ascrerr ****      � ****��  ��   L     o    ���� 0 errmsg errMsg   l     ��������  ��  ��    !"! i   , /#$# I      ��%���� 00 createzipwithlocal7zip CreateZipWithLocal7Zip% &��& o      ���� 0 
zipcommand 
zipCommand��  ��  $ Q     '()' k    ** +,+ I   ��-��
�� .sysoexecTEXT���     TEXT- o    ���� 0 
zipcommand 
zipCommand��  , .��. L   	 // m   	 
00 �11  S u c c e s s��  ( R      ������
�� .ascrerr ****      � ****��  ��  ) L    22 o    ���� 0 errmsg errMsg" 343 l     ��������  ��  ��  4 565 i   0 3787 I      ��9���� <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiver9 :��: o      ���� 0 paramstring paramString��  ��  8 k     <;; <=< l     ��>?��  > q k Create a ZIP file of all the PDFs in the target folder. Makes it simpler for you to send them to your KTs.   ? �@@ �   C r e a t e   a   Z I P   f i l e   o f   a l l   t h e   P D F s   i n   t h e   t a r g e t   f o l d e r .   M a k e s   i t   s i m p l e r   f o r   y o u   t o   s e n d   t h e m   t o   y o u r   K T s .= ABA r     CDC I      ��E���� 0 splitstring SplitStringE FGF o    ���� 0 paramstring paramStringG H��H m    II �JJ  - , -��  ��  D J      KK LML o      ���� 0 savepath savePathM N��N o      ���� 0 zippath zipPath��  B O��O Q    <PQRP k    2SS TUT I   /��V��
�� .sysoexecTEXT���     TEXTV b    +WXW b    )YZY b    '[\[ b    #]^] b    !_`_ b    aba m    cc �dd  c db 1    ��
�� 
spac` n     efe 1     ��
�� 
strqf o    ���� 0 savepath savePath^ m   ! "gg �hh (   & &   / u s r / b i n / z i p   - j  \ n   # &iji 1   $ &��
�� 
strqj o   # $���� 0 zippath zipPathZ 1   ' (��
�� 
spacX m   ) *kk �ll 
 * . p d f��  U m��m L   0 2nn m   0 1oo �pp  S u c c e s s��  Q R      ������
�� .ascrerr ****      � ****��  ��  R L   : <qq o   : ;���� 0 errmsg errMsg��  6 rsr l     ��������  ��  ��  s tut i   4 7vwv I      ��x���� 0 
deletefile 
DeleteFilex y��y o      ���� 0 filepath filePath��  ��  w k     zz {|{ l     ��}~��  } M GSelf-explanatory. This will delete the target file, skipping the Trash.   ~ � � S e l f - e x p l a n a t o r y .   T h i s   w i l l   d e l e t e   t h e   t a r g e t   f i l e ,   s k i p p i n g   t h e   T r a s h .| ��� l      ������  � � � The value of filePath passed to this function is always carefully considered
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
�� boovfals��  u ��� l     ��������  ��  ��  � ��� i   8 ;��� I      ������� "0 doesbundleexist DoesBundleExist� ���� o      ���� 0 
bundlepath 
bundlePath��  ��  � k     �� ��� l     ������  � D > Used to check if the Dialog Toolkit Plus script bundle exists   � ��� |   U s e d   t o   c h e c k   i f   t h e   D i a l o g   T o o l k i t   P l u s   s c r i p t   b u n d l e   e x i s t s� ���� O    ��� L    �� l   ������ I   ���~
� .coredoexnull���     ****� 4    �}�
�} 
ditm� o    �|�| 0 
bundlepath 
bundlePath�~  ��  ��  � m     ���                                                                                  sevs  alis    N  macOS                      �z2[BD ����System Events.app                                              �����z2[        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  ��  � ��� l     �{�z�y�{  �z  �y  � ��� i   < ?��� I      �x��w�x 0 doesfileexist DoesFileExist� ��v� o      �u�u 0 filepath filePath�v  �w  � k     �� ��� l     �t���t  �   Self-explanatory   � ��� "   S e l f - e x p l a n a t o r y� ��s� O    ��� L    �� F    ��� l   ��r�q� I   �p��o
�p .coredoexnull���     ****� 4    �n�
�n 
ditm� o    �m�m 0 filepath filePath�o  �r  �q  � =    ��� n    ��� m    �l
�l 
pcls� 4    �k�
�k 
ditm� o    �j�j 0 filepath filePath� m    �i
�i 
file� m     ���                                                                                  sevs  alis    N  macOS                      �z2[BD ����System Events.app                                              �����z2[        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  �s  � ��� l     �h�g�f�h  �g  �f  � ��� i   @ C��� I      �e��d�e 0 downloadfile DownloadFile� ��c� o      �b�b 0 paramstring paramString�c  �d  � k     B�� ��� l     �a���a  � Z T Self-explanatory. The value of fileURL is the internet address to the desired file.   � ��� �   S e l f - e x p l a n a t o r y .   T h e   v a l u e   o f   f i l e U R L   i s   t h e   i n t e r n e t   a d d r e s s   t o   t h e   d e s i r e d   f i l e .� ��� r     ��� I      �`��_�` 0 splitstring SplitString� ��� o    �^�^ 0 paramstring paramString� ��]� m    �� ���  - , -�]  �_  � J      �� ��� o      �\�\ "0 destinationpath destinationPath� ��[� o      �Z�Z 0 fileurl fileURL�[  � ��Y� Q    B���� k    .�� ��� I   +�X��W
�X .sysoexecTEXT���     TEXT� b    '��� b    #��� b    !��� b    ��� m    �� ���  c u r l   - L   - o� 1    �V
�V 
spac� l    ��U�T� n     ��� 1     �S
�S 
strq� o    �R�R "0 destinationpath destinationPath�U  �T  � 1   ! "�Q
�Q 
spac� l  # &��P�O� n   # &��� 1   $ &�N
�N 
strq� o   # $�M�M 0 fileurl fileURL�P  �O  �W  �  �L  L   , . m   , -�K
�K boovtrue�L  � R      �J�I�H
�J .ascrerr ****      � ****�I  �H  � k   6 B  I  6 ?�G�F
�G .sysodlogaskr        TEXT b   6 ; b   6 9	 m   6 7

 � . E r r o r   d o w n l o a d i n g   f i l e :	 1   7 8�E
�E 
spac o   9 :�D�D 0 fileurl fileURL�F   �C L   @ B m   @ A�B
�B boovfals�C  �Y  �  l     �A�@�?�A  �@  �?    i   D G I      �>�=�> 0 findsignature FindSignature �< o      �;�; 0 signaturepath signaturePath�<  �=   k     3  l     �:�:   m g If your signature isn't embedded in the Excel file, it will try to find an external JPG or PNG version    � �   I f   y o u r   s i g n a t u r e   i s n ' t   e m b e d d e d   i n   t h e   E x c e l   f i l e ,   i t   w i l l   t r y   t o   f i n d   a n   e x t e r n a l   J P G   o r   P N G   v e r s i o n �9 Q     3 Z    ) !"#  I    �8$�7�8 0 doesfileexist DoesFileExist$ %�6% b    &'& o    �5�5 0 signaturepath signaturePath' m    (( �))  m y S i g n a t u r e . p n g�6  �7  ! L    ** b    +,+ o    �4�4 0 signaturepath signaturePath, m    -- �..  m y S i g n a t u r e . p n g" /0/ I    �31�2�3 0 doesfileexist DoesFileExist1 2�12 b    343 o    �0�0 0 signaturepath signaturePath4 m    55 �66  m y S i g n a t u r e . j p g�1  �2  0 7�/7 L     $88 b     #9:9 o     !�.�. 0 signaturepath signaturePath: m   ! ";; �<<  m y S i g n a t u r e . p n g�/  # L   ' )== m   ' (>> �??   R      �-�,�+
�- .ascrerr ****      � ****�,  �+   L   1 3@@ m   1 2AA �BB  �9   CDC l     �*�)�(�*  �)  �(  D EFE i   H KGHG I      �'I�&�' 0 
renamefile 
RenameFileI J�%J o      �$�$ 0 paramstring paramString�%  �&  H k     DKK LML l     �#NO�#  N z t This pulls double duty for renaming a file or moving it to a new location. (It's the same process to the computer.)   O �PP �   T h i s   p u l l s   d o u b l e   d u t y   f o r   r e n a m i n g   a   f i l e   o r   m o v i n g   i t   t o   a   n e w   l o c a t i o n .   ( I t ' s   t h e   s a m e   p r o c e s s   t o   t h e   c o m p u t e r . )M QRQ r     STS I      �"U�!�" 0 splitstring SplitStringU VWV o    � �  0 paramstring paramStringW X�X m    YY �ZZ  - , -�  �!  T J      [[ \]\ o      �� 0 
targetfile 
targetFile] ^�^ o      �� 0 newfilename newFilename�  R _`_ r    aba n    cdc 1    �
� 
strqd n    efe 1    �
� 
psxpf o    �� 0 
targetfile 
targetFileb o      �� 0 
targetfile 
targetFile` ghg r    &iji n    $klk 1   " $�
� 
strql n    "mnm 1     "�
� 
psxpn o     �� 0 newfilename newFilenamej o      �� 0 newfilename newFilenameh o�o Q   ' Dpqrp k   * :ss tut I  * 7�v�
� .sysoexecTEXT���     TEXTv b   * 3wxw b   * 1yzy b   * /{|{ b   * -}~} m   * + ��� 
 m v   - f~ 1   + ,�
� 
spac| o   - .�� 0 
targetfile 
targetFilez 1   / 0�
� 
spacx o   1 2�� 0 newfilename newFilename�  u ��� L   8 :�� m   8 9�
� boovtrue�  q R      �
�	�
�
 .ascrerr ****      � ****�	  �  r L   B D�� m   B C�
� boovfals�  F ��� l     ����  �  �  � ��� l     ����  �   Folder Manipulation   � ��� (   F o l d e r   M a n i p u l a t i o n� ��� l     ��� �  �  �   � ��� i   L O��� I      ������� 0 clearfolder ClearFolder� ���� o      ���� 0 foldertoempty folderToEmpty��  ��  � k     ?�� ��� l     ������  � h b Empties the target folder, but only of DOCX, PDF, and ZIP files. This folder will not be deleted.   � ��� �   E m p t i e s   t h e   t a r g e t   f o l d e r ,   b u t   o n l y   o f   D O C X ,   P D F ,   a n d   Z I P   f i l e s .   T h i s   f o l d e r   w i l l   n o t   b e   d e l e t e d .� ���� Q     ?���� k    5�� ��� I   �����
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
spac� m   , -�� ��� < - t y p e   f   - n a m e   ' * . d o c x '   - d e l e t e��  � ���� L   3 5�� m   3 4��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   = ?�� m   = >��
�� boovfals��  � ��� l     ��������  ��  ��  � ��� i   P S��� I      ������� .0 clearpdfsafterzipping ClearPDFsAfterZipping� ���� o      ���� 0 foldertoempty folderToEmpty��  ��  � Q     ���� k    �� ��� I   �����
�� .sysoexecTEXT���     TEXT� b    ��� b    ��� b    
��� b    ��� m    �� ���  f i n d� 1    ��
�� 
spac� l   	������ n    	��� 1    	��
�� 
strq� o    ���� 0 foldertoempty folderToEmpty��  ��  � 1   
 ��
�� 
spac� m    �� ��� : - t y p e   f   - n a m e   ' * . p d f '   - d e l e t e��  � ���� L    �� m    ��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L    �� m    ��
�� boovfals� ��� l     ��������  ��  ��  � ��� i   T W��� I      ������� 0 
copyfolder 
CopyFolder� ���� o      ���� 0 
folderpath 
folderPath��  ��  � k     8    l     ����   o i Self-explanatory. Copy a folder (or bundle) from place A to place B. The original file will still exist.    � �   S e l f - e x p l a n a t o r y .   C o p y   a   f o l d e r   ( o r   b u n d l e )   f r o m   p l a c e   A   t o   p l a c e   B .   T h e   o r i g i n a l   f i l e   w i l l   s t i l l   e x i s t .  r     	 I      ��
���� 0 splitstring SplitString
  o    ���� 0 
folderpath 
folderPath �� m     �  - , -��  ��  	 J        o      ���� 0 targetfolder targetFolder �� o      ���� &0 destinationfolder destinationFolder��   �� Q    8 k    .  I   +����
�� .sysoexecTEXT���     TEXT b    ' b    # b    ! !  b    "#" m    $$ �%%  c p   - R f# 1    ��
�� 
spac! l    &����& n     '(' 1     ��
�� 
strq( o    ���� 0 targetfolder targetFolder��  ��   1   ! "��
�� 
spac l  # &)����) n   # &*+* 1   $ &��
�� 
strq+ o   # $���� &0 destinationfolder destinationFolder��  ��  ��   ,��, L   , .-- m   , -��
�� boovtrue��   R      ������
�� .ascrerr ****      � ****��  ��   L   6 8.. m   6 7��
�� boovfals��  � /0/ l     ��������  ��  ��  0 121 i   X [343 I      ��5���� 0 createfolder CreateFolder5 6��6 o      ���� 0 
folderpath 
folderPath��  ��  4 k     77 898 l     ��:;��  : \ V Self-explanatory. Needed for creating the folder for where the reports will be saved.   ; �<< �   S e l f - e x p l a n a t o r y .   N e e d e d   f o r   c r e a t i n g   t h e   f o l d e r   f o r   w h e r e   t h e   r e p o r t s   w i l l   b e   s a v e d .9 =��= Q     >?@> k    AA BCB I   ��D��
�� .sysoexecTEXT���     TEXTD b    
EFE b    GHG m    II �JJ  m k d i r   - pH 1    ��
�� 
spacF l   	K����K n    	LML 1    	��
�� 
strqM o    ���� 0 
folderpath 
folderPath��  ��  ��  C N��N L    OO m    ��
�� boovtrue��  ? R      ������
�� .ascrerr ****      � ****��  ��  @ L    PP m    ��
�� boovfals��  2 QRQ l     ��������  ��  ��  R STS i   \ _UVU I      ��W���� 0 deletefolder DeleteFolderW X��X o      ���� 0 
folderpath 
folderPath��  ��  V k     YY Z[Z l     ��\]��  \ c ] Self-explanatory. Same as with DeleteFile, extra security checks will likely be added later.   ] �^^ �   S e l f - e x p l a n a t o r y .   S a m e   a s   w i t h   D e l e t e F i l e ,   e x t r a   s e c u r i t y   c h e c k s   w i l l   l i k e l y   b e   a d d e d   l a t e r .[ _��_ Q     `ab` k    cc ded I   ��f��
�� .sysoexecTEXT���     TEXTf b    
ghg b    iji m    kk �ll  r m   - r fj 1    ��
�� 
spach l   	m����m n    	non 1    	��
�� 
strqo o    �� 0 
folderpath 
folderPath��  ��  ��  e p�~p L    qq m    �}
�} boovtrue�~  a R      �|�{�z
�| .ascrerr ****      � ****�{  �z  b L    rr m    �y
�y boovfals��  T sts l     �x�w�v�x  �w  �v  t uvu i   ` cwxw I      �uy�t�u "0 doesfolderexist DoesFolderExisty z�sz o      �r�r 0 
folderpath 
folderPath�s  �t  x k     {{ |}| l     �q~�q  ~   Self-explanatory    ��� "   S e l f - e x p l a n a t o r y} ��p� O    ��� L    �� F    ��� l   ��o�n� I   �m��l
�m .coredoexnull���     ****� 4    �k�
�k 
ditm� o    �j�j 0 
folderpath 
folderPath�l  �o  �n  � =    ��� n    ��� m    �i
�i 
pcls� 4    �h�
�h 
ditm� o    �g�g 0 
folderpath 
folderPath� m    �f
�f 
cfol� m     ���                                                                                  sevs  alis    N  macOS                      �z2[BD ����System Events.app                                              �����z2[        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  �p  v ��� l     �e�d�c�e  �d  �c  � ��� l     �b���b  �   Dialog Boxes   � ���    D i a l o g   B o x e s� ��� l     �a�`�_�a  �`  �_  � ��� i   d g��� I      �^��]�^ 80 installdialogdisplayscript InstallDialogDisplayScript� ��\� o      �[�[ 0 paramstring paramString�\  �]  � k     �� ��� r     ��� b     	��� n     ��� 1    �Z
�Z 
psxp� l    ��Y�X� I    �W��V
�W .earsffdralis        afdr� m     �U
�U afdrcusr�V  �Y  �X  � m    �� ��� � L i b r a r y / A p p l i c a t i o n   S c r i p t s / c o m . m i c r o s o f t . E x c e l / D i a l o g D i s p l a y . s c p t� o      �T�T 0 
scriptpath 
scriptPath� ��� r    ��� m    �� ��� � h t t p s : / / r a w . g i t h u b u s e r c o n t e n t . c o m / p a p e r c u t t e r 0 3 2 4 / S p e a k i n g E v a l s / m a i n / D i a l o g D i s p l a y . s c p t� o      �S�S 0 downloadurl downloadURL� ��� l   �R�Q�P�R  �Q  �P  � ��� l   �O���O  � A ; If an existing version is not found, download a fresh copy   � ��� v   I f   a n   e x i s t i n g   v e r s i o n   i s   n o t   f o u n d ,   d o w n l o a d   a   f r e s h   c o p y� ��� l   �N���N  � e _ Skip this first check until a full update function can be designed. For now, install each time   � ��� �   S k i p   t h i s   f i r s t   c h e c k   u n t i l   a   f u l l   u p d a t e   f u n c t i o n   c a n   b e   d e s i g n e d .   F o r   n o w ,   i n s t a l l   e a c h   t i m e� ��� l   �M���M  � 4 . if DoesFileExist(scriptPath) then return true   � ��� \   i f   D o e s F i l e E x i s t ( s c r i p t P a t h )   t h e n   r e t u r n   t r u e� ��L� L    �� I    �K��J�K 0 downloadfile DownloadFile� ��I� b    ��� b    ��� o    �H�H 0 
scriptpath 
scriptPath� m    �� ���  - , -� o    �G�G 0 downloadurl downloadURL�I  �J  �L  � ��� l     �F�E�D�F  �E  �D  � ��� i   h k��� I      �C��B�C >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder� ��A� o      �@�@ 0 paramstring paramString�A  �B  � k     ]�� ��� r     ��� b     	��� n     ��� 1    �?
�? 
psxp� l    ��>�=� I    �<��;
�< .earsffdralis        afdr� m     �:
�: afdrcusr�;  �>  �=  � m    �� ��� 0 L i b r a r y / S c r i p t   L i b r a r i e s� o      �9�9 .0 scriptlibrariesfolder scriptLibrariesFolder� ��� l   �8�7�6�8  �7  �6  � ��5� Z    ]���4�� I    �3��2�3 "0 doesfolderexist DoesFolderExist� ��1� o    �0�0 .0 scriptlibrariesfolder scriptLibrariesFolder�1  �2  � L    �� o    �/�/ .0 scriptlibrariesfolder scriptLibrariesFolder�4  � Q    ]���� k    Q�� ��� l   �.���.  � m g ~/Library is typically a read-only folder, so I need to requst your password to create the need folder   � ��� �   ~ / L i b r a r y   i s   t y p i c a l l y   a   r e a d - o n l y   f o l d e r ,   s o   I   n e e d   t o   r e q u s t   y o u r   p a s s w o r d   t o   c r e a t e   t h e   n e e d   f o l d e r� ��� I   *�-��
�- .sysoexecTEXT���     TEXT� b    $��� b     ��� m    �� ���  m k d i r   - p� 1    �,
�, 
spac� n     #��� 1   ! #�+
�+ 
strq� o     !�*�* .0 scriptlibrariesfolder scriptLibrariesFolder� �) �(
�) 
badm  m   % &�'
�' boovtrue�(  �  l  + +�&�&   %  Set your username as the owner    � >   S e t   y o u r   u s e r n a m e   a s   t h e   o w n e r  I  + B�%	
�% .sysoexecTEXT���     TEXT b   + <

 b   + 8 b   + 6 m   + , �  c h o w n   n   , 5 1   3 5�$
�$ 
strq l  , 3�#�" n   , 3 1   1 3�!
�! 
sisn l  , 1� � I  , 1���
� .sysosigtsirr   ��� null�  �  �   �  �#  �"   1   6 7�
� 
spac n   8 ; 1   9 ;�
� 
strq o   8 9�� .0 scriptlibrariesfolder scriptLibrariesFolder	 ��
� 
badm m   = >�
� boovtrue�    l  C C��   5 / Give your username READ and WRITE permissions.    � ^   G i v e   y o u r   u s e r n a m e   R E A D   a n d   W R I T E   p e r m i s s i o n s .  !  I  C N�"#
� .sysoexecTEXT���     TEXT" b   C H$%$ m   C D&& �''  c h m o d   u + r w  % n   D G()( 1   E G�
� 
strq) o   D E�� .0 scriptlibrariesfolder scriptLibrariesFolder# �*�
� 
badm* m   I J�
� boovtrue�  ! +�+ L   O Q,, o   O P�� .0 scriptlibrariesfolder scriptLibrariesFolder�  � R      ���

� .ascrerr ****      � ****�  �
  � L   Y ]-- m   Y \.. �//  �5  � 010 l     �	���	  �  �  1 232 i   l o454 I      �6�� 40 installdialogtoolkitplus InstallDialogToolkitPlus6 7�7 o      �� "0 resourcesfolder resourcesFolder�  �  5 k     �88 9:9 r     ;<; m     == �>> � h t t p s : / / r a w . g i t h u b u s e r c o n t e n t . c o m / p a p e r c u t t e r 0 3 2 4 / S p e a k i n g E v a l s / m a i n / D i a l o g _ T o o l k i t . z i p< o      �� 0 downloadurl downloadURL: ?@? r    ABA b    CDC n    EFE 1   	 �
� 
psxpF l   	G� ��G I   	��H��
�� .earsffdralis        afdrH m    ��
�� afdrcusr��  �   ��  D m    II �JJ 0 L i b r a r y / S c r i p t   L i b r a r i e sB o      ���� .0 scriptlibrariesfolder scriptLibrariesFolder@ KLK r    MNM m    OO �PP 4 / D i a l o g   T o o l k i t   P l u s . s c p t dN o      ���� $0 dialogbundlename dialogBundleNameL QRQ r    STS b    UVU o    ���� .0 scriptlibrariesfolder scriptLibrariesFolderV o    ���� $0 dialogbundlename dialogBundleNameT o      ���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundleR WXW r    YZY b    [\[ o    ���� "0 resourcesfolder resourcesFolder\ m    ]] �^^ & / D i a l o g _ T o o l k i t . z i pZ o      ���� 0 zipfilepath zipFilePathX _`_ r     %aba b     #cdc o     !���� "0 resourcesfolder resourcesFolderd m   ! "ee �ff $ / d i a l o g T o o l k i t T e m pb o      ���� &0 zipextractionpath zipExtractionPath` ghg l  & &��������  ��  ��  h iji l  & &��kl��  k 0 * Initial check to see if already installed   l �mm T   I n i t i a l   c h e c k   t o   s e e   i f   a l r e a d y   i n s t a l l e dj non Z  & 5pq����p I   & ,��r���� "0 doesbundleexist DoesBundleExistr s��s o   ' (���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  q L   / 1tt m   / 0��
�� boovtrue��  ��  o uvu l  6 6��������  ��  ��  v wxw l  6 6��yz��  y 3 - Ensure resources folder exists for later use   z �{{ Z   E n s u r e   r e s o u r c e s   f o l d e r   e x i s t s   f o r   l a t e r   u s ex |}| Z   6 W~����~ H   6 =�� I   6 <������� "0 doesfolderexist DoesFolderExist� ���� o   7 8���� "0 resourcesfolder resourcesFolder��  ��   Q   @ S���� I   C I������� 0 createfolder CreateFolder� ���� o   D E���� "0 resourcesfolder resourcesFolder��  ��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   Q S�� m   Q R��
�� boovfals��  ��  } ��� l  X X��������  ��  ��  � ��� l  X X������  � G A Check for a local copy and move it to the needed folder if found   � ��� �   C h e c k   f o r   a   l o c a l   c o p y   a n d   m o v e   i t   t o   t h e   n e e d e d   f o l d e r   i f   f o u n d� ��� Z   X |������� I   X `������� "0 doesbundleexist DoesBundleExist� ���� b   Y \��� o   Y Z���� "0 resourcesfolder resourcesFolder� o   Z [���� $0 dialogbundlename dialogBundleName��  ��  � Z   c x������� I   c o������� 0 
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
CopyFolder� ���� b   � ���� b   � ���� b   � ���� b   � ���� b   � ���� o   � ����� &0 zipextractionpath zipExtractionPath� m   � ��� ���  / D i a l o g _ T o o l k i t� o   � ����� $0 dialogbundlename dialogBundleName� m   � ��� ���  - , -� o   � ����� "0 resourcesfolder resourcesFolder� o   � ����� $0 dialogbundlename dialogBundleName��  ��  � ��� l  � �������  � ; 5 ...and copy the script bundle to the required folder   � ��� j   . . . a n d   c o p y   t h e   s c r i p t   b u n d l e   t o   t h e   r e q u i r e d   f o l d e r� ���� I   � �������� 0 
copyfolder 
CopyFolder� ���� b   � ���� b   � ���� b   � ���� b   � ���� o   � ����� &0 zipextractionpath zipExtractionPath� m   � ��� ���  / D i a l o g _ T o o l k i t� o   � ����� $0 dialogbundlename dialogBundleName� m   � ��� ���  - , -� o   � ����� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  ��  � R      ������
�� .ascrerr ****      � ****��  ��  ��  ��  ��  �    l  � ���������  ��  ��    l  � �����   D > Remove unneeded files and folders created during this process    � |   R e m o v e   u n n e e d e d   f i l e s   a n d   f o l d e r s   c r e a t e d   d u r i n g   t h i s   p r o c e s s  I   � ���	���� 0 
deletefile 
DeleteFile	 
��
 o   � ����� 0 zipfilepath zipFilePath��  ��    I   � ������� 0 deletefolder DeleteFolder �� o   � ����� &0 zipextractionpath zipExtractionPath��  ��    l  � ���������  ��  ��    l  � �����   V P One final check to verify installation was successful and return true if it was    � �   O n e   f i n a l   c h e c k   t o   v e r i f y   i n s t a l l a t i o n   w a s   s u c c e s s f u l   a n d   r e t u r n   t r u e   i f   i t   w a s �� L   � � I   � ������� "0 doesbundleexist DoesBundleExist �� o   � ����� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  ��  3  l     �������  ��  �    i   p s I      �~ �}�~ 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus  !�|! o      �{�{ "0 resourcesfolder resourcesFolder�|  �}   k     U"" #$# r     %&% b     	'(' n     )*) 1    �z
�z 
psxp* l    +�y�x+ I    �w,�v
�w .earsffdralis        afdr, m     �u
�u afdrcusr�v  �y  �x  ( m    -- �.. d L i b r a r y / S c r i p t   L i b r a r i e s / D i a l o g   T o o l k i t   P l u s . s c p t d& o      �t�t 20 dialogtoolkitplusbundle dialogToolkitPlusBundle$ /0/ r    121 b    343 o    �s�s "0 resourcesfolder resourcesFolder4 m    55 �66 4 / D i a l o g   T o o l k i t   P l u s . s c p t d2 o      �r�r 0 	localcopy 	localCopy0 787 l   �q�p�o�q  �p  �o  8 9:9 Z    R;<�n=; I    �m>�l�m "0 doesbundleexist DoesBundleExist> ?�k? o    �j�j 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�k  �l  < Q    L@AB@ k    ACC DED Z   6FG�i�hF H    %HH I    $�gI�f�g "0 doesbundleexist DoesBundleExistI J�eJ o     �d�d 0 	localcopy 	localCopy�e  �f  G I   ( 2�cK�b�c 0 
copyfolder 
CopyFolderK L�aL b   ) .MNM b   ) ,OPO o   ) *�`�` 20 dialogtoolkitplusbundle dialogToolkitPlusBundleP m   * +QQ �RR  - , -N o   , -�_�_ 0 	localcopy 	localCopy�a  �b  �i  �h  E STS I   7 =�^U�]�^ 0 deletefolder DeleteFolderU V�\V o   8 9�[�[ 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�\  �]  T W�ZW r   > AXYX m   > ?�Y
�Y boovtrueY o      �X�X 0 removalresult removalResult�Z  A R      �W�V�U
�W .ascrerr ****      � ****�V  �U  B r   I LZ[Z m   I J�T
�T boovfals[ o      �S�S 0 removalresult removalResult�n  = r   O R\]\ m   O P�R
�R boovtrue] o      �Q�Q 0 removalresult removalResult: ^_^ l  S S�P�O�N�P  �O  �N  _ `�M` L   S Uaa o   S T�L�L 0 removalresult removalResult�M   b�Kb l     �J�I�H�J  �I  �H  �K       �Gcdefghijklmnopqrstuvwxyz{|}~��G  c �F�E�D�C�B�A�@�?�>�=�<�;�:�9�8�7�6�5�4�3�2�1�0�/�.�-�,�+�*�F 00 getscriptversionnumber GetScriptVersionNumber�E "0 getmacosversion GetMacOSVersion�D 80 checkaccessibilitysettings CheckAccessibilitySettings�C 0 splitstring SplitString�B "0 loadapplication LoadApplication�A 0 isapploaded IsAppLoaded�@ 0 	closeword 	CloseWord�? .0 changefilepermissions ChangeFilePermissions�> $0 comparemd5hashes CompareMD5Hashes�= 0 copyfile CopyFile�< (0 createzipwithditto CreateZipWithDitto�; 00 createzipwithlocal7zip CreateZipWithLocal7Zip�: <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiver�9 0 
deletefile 
DeleteFile�8 "0 doesbundleexist DoesBundleExist�7 0 doesfileexist DoesFileExist�6 0 downloadfile DownloadFile�5 0 findsignature FindSignature�4 0 
renamefile 
RenameFile�3 0 clearfolder ClearFolder�2 .0 clearpdfsafterzipping ClearPDFsAfterZipping�1 0 
copyfolder 
CopyFolder�0 0 createfolder CreateFolder�/ 0 deletefolder DeleteFolder�. "0 doesfolderexist DoesFolderExist�- 80 installdialogdisplayscript InstallDialogDisplayScript�, >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder�+ 40 installdialogtoolkitplus InstallDialogToolkitPlus�* 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlusd �) �(�'���&�) 00 getscriptversionnumber GetScriptVersionNumber�( �%��% �  �$�$ 0 paramstring paramString�'  � �#�# 0 paramstring paramString� �"�" 4�s�& �e �! &� �����! "0 getmacosversion GetMacOSVersion�  ��� �  �� 0 paramstring paramString�  � ��� 0 paramstring paramString� 0 	osversion 	osVersion�  8���
� .sysoexecTEXT���     TEXT�  �  �  �j E�O�W X  hf � A������ 80 checkaccessibilitysettings CheckAccessibilitySettings� ��� �  �� 0 
apptocheck 
appToCheck�  � ��� 0 
apptocheck 
appToCheck� ,0 accessibilityenabled accessibilityEnabled�  n������
�	���
� 
prcs
� 
pnam�  
� 
pvis
� 
pcap
�
 
uiel
�	 
enaB
� 
bool�  �  � 6 -� %*�-�,�[�,\Ze81�	 *�/�-�,E�&E�O�UW 	X 	 
fg � |������ 0 splitstring SplitString� ��� �  � ���  &0 passedparamstring passedParamString�� (0 parameterseparator parameterSeparator�  � ���������� &0 passedparamstring passedParamString�� (0 parameterseparator parameterSeparator�� 00 oldtextitemsdelimiters oldTextItemsDelimiters�� *0 separatedparameters separatedParameters� ������
�� 
ascr
�� 
txdl
�� 
citm�  � *�,E�O�*�,FO��-E�O�*�,FUO�h �� ����������� "0 loadapplication LoadApplication�� ����� �  ���� 0 appname appName��  � �������� 0 appname appName�� 0 errmsg errMsg�� 0 errnum errNum� 	���� ���� ��� � �
�� 
capp
�� .miscactvnull��� ��� null�� 0 errmsg errMsg� ������
�� 
errn�� 0 errnum errNum��  
�� 
spac�� * *�/ *j UO�W X  ��%�%�%�%�%�%i �� ����������� 0 isapploaded IsAppLoaded�� ����� �  ���� 0 appname appName��  � ���������� 0 appname appName�� 0 
loadresult 
loadResult�� 0 errmsg errMsg�� 0 errnum errNum� ������ ����
�� 
prcs
�� 
pnam
�� 
spac�� 0 errmsg errMsg� ������
�� 
errn�� 0 errnum errNum��  �� ; (� *�-�,� ��%�%E�Y 	��%�%E�UO�W X  �%�%�%�%�%j ��#���������� 0 	closeword 	CloseWord�� ����� �  ���� 0 paramstring paramString��  � ������ 0 paramstring paramString�� 0 closeresult closeResult� P����=D��HL����R
�� 
prcs
�� 
pnam
�� .aevtquitnull��� ��� null��  ��  �� 4 +� #*�-�,� � *j UO�E�Y �E�O�UW 	X  	�k ��`���������� .0 changefilepermissions ChangeFilePermissions�� ����� �  ���� 0 paramstring paramString��  � �������� 0 paramstring paramString��  0 newpermissions newPermissions�� 0 filepath filePath� 
l����~������������� 0 splitstring SplitString
�� 
cobj
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� C*��l+ E[�k/E�Z[�l/E�ZO #��%��,%j O��%�%�%��,%j OeW 	X  	fl ������������� $0 comparemd5hashes CompareMD5Hashes�� ����� �  ���� 0 paramstring paramString��  � ���������� 0 paramstring paramString�� 0 filepath filePath�� 0 	validhash 	validHash�� 0 checkresult checkResult� 
�������������������� 0 splitstring SplitString
�� 
cobj�� 0 doesfileexist DoesFileExist
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� H*��l+ E[�k/E�Z[�l/E�ZO*�k+  fY hO ��%��,%j E�O�� W 	X  	fm ������������� 0 copyfile CopyFile�� ����� �  ���� 0 	filepaths 	filePaths��  � �������� 0 	filepaths 	filePaths�� 0 
targetfile 
targetFile�� "0 destinationfile destinationFile� 	����� ������������ 0 splitstring SplitString
�� 
cobj
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� 9*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW 	X  fn ������������ (0 createzipwithditto CreateZipWithDitto�� ����� �  ���� 0 
zipcommand 
zipCommand��  � ������ 0 
zipcommand 
zipCommand�� 0 errmsg errMsg� ������
�� .sysoexecTEXT���     TEXT��  ��  ��  �j  O�W 	X  �o ��$���������� 00 createzipwithlocal7zip CreateZipWithLocal7Zip�� ����� �  ���� 0 
zipcommand 
zipCommand��  � ������ 0 
zipcommand 
zipCommand�� 0 errmsg errMsg� ��0����
�� .sysoexecTEXT���     TEXT��  ��  ��  �j  O�W 	X  �p ��8������~�� <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiver�� �}��} �  �|�| 0 paramstring paramString�  � �{�z�y�x�{ 0 paramstring paramString�z 0 savepath savePath�y 0 zippath zipPath�x 0 errmsg errMsg� I�w�vc�u�tgk�so�r�q�w 0 splitstring SplitString
�v 
cobj
�u 
spac
�t 
strq
�s .sysoexecTEXT���     TEXT�r  �q  �~ =*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%�%�%j O�W 	X 
 �q �pw�o�n���m�p 0 
deletefile 
DeleteFile�o �l��l �  �k�k 0 filepath filePath�n  � �j�j 0 filepath filePath� ��i�h�g�f�e
�i 
spac
�h 
strq
�g .sysoexecTEXT���     TEXT�f  �e  �m  ��%��,%j OeW 	X  fr �d��c�b���a�d "0 doesbundleexist DoesBundleExist�c �`��` �  �_�_ 0 
bundlepath 
bundlePath�b  � �^�^ 0 
bundlepath 
bundlePath� ��]�\
�] 
ditm
�\ .coredoexnull���     ****�a � *�/j Us �[��Z�Y���X�[ 0 doesfileexist DoesFileExist�Z �W��W �  �V�V 0 filepath filePath�Y  � �U�U 0 filepath filePath� ��T�S�R�Q�P
�T 
ditm
�S .coredoexnull���     ****
�R 
pcls
�Q 
file
�P 
bool�X � *�/j 	 *�/�,� �&Ut �O��N�M���L�O 0 downloadfile DownloadFile�N �K��K �  �J�J 0 paramstring paramString�M  � �I�H�G�I 0 paramstring paramString�H "0 destinationpath destinationPath�G 0 fileurl fileURL� ��F�E��D�C�B�A�@
�?�F 0 splitstring SplitString
�E 
cobj
�D 
spac
�C 
strq
�B .sysoexecTEXT���     TEXT�A  �@  
�? .sysodlogaskr        TEXT�L C*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW X  ��%�%j 
Ofu �>�=�<���;�> 0 findsignature FindSignature�= �:��: �  �9�9 0 signaturepath signaturePath�<  � �8�8 0 signaturepath signaturePath� 	(�7-5;>�6�5A�7 0 doesfileexist DoesFileExist�6  �5  �; 4 +*��%k+  	��%Y *��%k+  	��%Y �W 	X  �v �4H�3�2���1�4 0 
renamefile 
RenameFile�3 �0��0 �  �/�/ 0 paramstring paramString�2  � �.�-�,�. 0 paramstring paramString�- 0 
targetfile 
targetFile�, 0 newfilename newFilename� 
Y�+�*�)�(�'�&�%�$�+ 0 splitstring SplitString
�* 
cobj
�) 
psxp
�( 
strq
�' 
spac
�& .sysoexecTEXT���     TEXT�%  �$  �1 E*��l+ E[�k/E�Z[�l/E�ZO��,�,E�O��,�,E�O ��%�%�%�%j OeW 	X  	fw �#��"�!��� �# 0 clearfolder ClearFolder�" ��� �  �� 0 foldertoempty folderToEmpty�!  � �� 0 foldertoempty folderToEmpty� �����������
� 
spac
� 
strq
� .sysoexecTEXT���     TEXT�  �  �  @ 7��%��,%�%�%j O��%��,%�%�%j O��%��,%�%�%j OeW 	X 	 
fx �������� .0 clearpdfsafterzipping ClearPDFsAfterZipping� ��� �  �� 0 foldertoempty folderToEmpty�  � �� 0 foldertoempty folderToEmpty� �������
� 
spac
� 
strq
� .sysoexecTEXT���     TEXT�  �  �   ��%��,%�%�%j OeW 	X  fy ���
�	���� 0 
copyfolder 
CopyFolder�
 ��� �  �� 0 
folderpath 
folderPath�	  � ���� 0 
folderpath 
folderPath� 0 targetfolder targetFolder� &0 destinationfolder destinationFolder� 	��$� ��������� 0 splitstring SplitString
� 
cobj
�  
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  � 9*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW 	X  fz ��4���������� 0 createfolder CreateFolder�� ����� �  ���� 0 
folderpath 
folderPath��  � ���� 0 
folderpath 
folderPath� I����������
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  ��  ��%��,%j OeW 	X  f{ ��V���������� 0 deletefolder DeleteFolder�� ����� �  ���� 0 
folderpath 
folderPath��  � ���� 0 
folderpath 
folderPath� k����������
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  ��  ��%��,%j OeW 	X  f| ��x���������� "0 doesfolderexist DoesFolderExist�� ����� �  ���� 0 
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
bool�� � *�/j 	 *�/�,� �&U} ������������� 80 installdialogdisplayscript InstallDialogDisplayScript�� ����� �  ���� 0 paramstring paramString��  � �������� 0 paramstring paramString�� 0 
scriptpath 
scriptPath�� 0 downloadurl downloadURL� �����������
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp�� 0 downloadfile DownloadFile�� �j �,�%E�O�E�O*��%�%k+ ~ ������������� >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder�� ����� �  ���� 0 paramstring paramString��  � ������ 0 paramstring paramString�� .0 scriptlibrariesfolder scriptLibrariesFolder� ����������������������&����.
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
sisn��  ��  �� ^�j �,�%E�O*�k+  �Y E 9��%��,%�el 	O�*j �,�,%�%��,%�el 	O���,%�el 	O�W X  a  ��5���������� 40 installdialogtoolkitplus InstallDialogToolkitPlus�� ����� �  ���� "0 resourcesfolder resourcesFolder��  � ���������������� "0 resourcesfolder resourcesFolder�� 0 downloadurl downloadURL�� .0 scriptlibrariesfolder scriptLibrariesFolder�� $0 dialogbundlename dialogBundleName�� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�� 0 zipfilepath zipFilePath�� &0 zipextractionpath zipExtractionPath� =������IO]e��������������������������������
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
W 	X  fY hO*��%k+  *��%�%�%k+  eY hY hO*��%�%k+  T Ha _ %�a ,%a %�a ,%j O*�a %�%a %�%�%k+ O*�a %�%a %�%k+ W X  hY hO*�k+ O*�k+ O*�k+ � ������������ 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus�� ����� �  ���� "0 resourcesfolder resourcesFolder��  � ���������� "0 resourcesfolder resourcesFolder�� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�� 0 	localcopy 	localCopy�� 0 removalresult removalResult� ������-5��Q��������
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp�� "0 doesbundleexist DoesBundleExist�� 0 
copyfolder 
CopyFolder�� 0 deletefolder DeleteFolder��  ��  �� V�j �,�%E�O��%E�O*�k+  6 (*�k+  *��%�%k+ Y hO*�k+ OeE�W 
X 	 
fE�Y eE�O� ascr  ��ޭ