FasdUAS 1.101.10   ��   ��    k             l      ��  ��    � |
Helper Scripts for the DYB Speaking Evaluations Excel spreadsheet

Version: 1.2.0
Build:   20250207
Warren Feltmate
� 2025
     � 	 	 � 
 H e l p e r   S c r i p t s   f o r   t h e   D Y B   S p e a k i n g   E v a l u a t i o n s   E x c e l   s p r e a d s h e e t 
 
 V e r s i o n :   1 . 2 . 0 
 B u i l d :       2 0 2 5 0 2 0 7 
 W a r r e n   F e l t m a t e 
 �   2 0 2 5 
   
  
 l     ��������  ��  ��        l     ��  ��      Environment Variables     �   ,   E n v i r o n m e n t   V a r i a b l e s      l     ��������  ��  ��        i         I      �� ���� 00 getscriptversionnumber GetScriptVersionNumber   ��  o      ���� 0 paramstring paramString��  ��    k            l     ��  ��    ? 9- Use build number to determine if an update is available     �   r -   U s e   b u i l d   n u m b e r   t o   d e t e r m i n e   i f   a n   u p d a t e   i s   a v a i l a b l e   ��  L          m     ���� 4�_��     ! " ! l     ��������  ��  ��   "  # $ # i     % & % I      �� '���� "0 getmacosversion GetMacOSVersion '  (�� ( o      ���� 0 paramstring paramString��  ��   & k      ) )  * + * l     �� , -��   , ` Z Not currently used, but could be helpful if there are issues with older versions of MacOS    - � . . �   N o t   c u r r e n t l y   u s e d ,   b u t   c o u l d   b e   h e l p f u l   i f   t h e r e   a r e   i s s u e s   w i t h   o l d e r   v e r s i o n s   o f   M a c O S +  /�� / Q      0 1�� 0 k     2 2  3 4 3 r    
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
�Z .aevtquitnull��� ��� null�Y  �X  C m    DD�                                                                                  MSWD  alis    4  macOS                      �z2[BD ����Microsoft Word.app                                             ����ㇾA        ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p    m a c O S  Applications/Microsoft Word.app   / ��  A E�WE r    FGF m    HH �II D W o r d   h a s   s u c c e s s f u l l y   b e e n   c l o s e d .G o      �V�V 0 closeresult closeResult�W  �_  7 r   " %JKJ m   " #LL �MM < W o r d   i s   n o t   c u r r e n t l y   r u n n i n g .K o      �U�U 0 closeresult closeResult4 N�TN L   & (OO o   & '�S�S 0 closeresult closeResult�T  1 m    PP�                                                                                  sevs  alis    N  macOS                      �z2[BD ����System Events.app                                              �����z2[        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  . R      �R�Q�P
�R .ascrerr ****      � ****�Q  �P  / L   1 3QQ m   1 2RR �SS P T h e r e   w a s   a n   e r r o r   t r y i n g   t o   c l o s e   W o r d .�`  ! TUT l     �O�N�M�O  �N  �M  U VWV l     �LXY�L  X   File Manipulation   Y �ZZ $   F i l e   M a n i p u l a t i o nW [\[ l     �K�J�I�K  �J  �I  \ ]^] i    _`_ I      �Ha�G�H $0 comparemd5hashes CompareMD5Hashesa b�Fb o      �E�E 0 paramstring paramString�F  �G  ` k     Gcc ded l     �Dfg�D  f b \ This will check the file integrity of the downloaded template against the known good value.   g �hh �   T h i s   w i l l   c h e c k   t h e   f i l e   i n t e g r i t y   o f   t h e   d o w n l o a d e d   t e m p l a t e   a g a i n s t   t h e   k n o w n   g o o d   v a l u e .e iji r     klk I      �Cm�B�C 0 splitstring SplitStringm non o    �A�A 0 paramstring paramStringo p�@p m    qq �rr  - , -�@  �B  l J      ss tut o      �?�? 0 filepath filePathu v�>v o      �=�= 0 	validhash 	validHash�>  j wxw l   �<�;�:�<  �;  �:  x yzy Z    '{|�9�8{ H    }} I    �7~�6�7 0 doesfileexist DoesFileExist~ �5 o    �4�4 0 filepath filePath�5  �6  | L   ! #�� m   ! "�3
�3 boovfals�9  �8  z ��� l  ( (�2�1�0�2  �1  �0  � ��/� Q   ( G���� k   + =�� ��� r   + 8��� l  + 6��.�-� I  + 6�,��+
�, .sysoexecTEXT���     TEXT� b   + 2��� b   + .��� m   + ,�� ���  m d 5   - q� 1   , -�*
�* 
spac� n   . 1��� 1   / 1�)
�) 
strq� o   . /�(�( 0 filepath filePath�+  �.  �-  � o      �'�' 0 checkresult checkResult� ��&� L   9 =�� =  9 <��� o   9 :�%�% 0 checkresult checkResult� o   : ;�$�$ 0 	validhash 	validHash�&  � R      �#�"�!
�# .ascrerr ****      � ****�"  �!  � L   E G�� m   E F� 
�  boovfals�/  ^ ��� l     ����  �  �  � ��� i     #��� I      ���� 0 copyfile CopyFile� ��� o      �� 0 	filepaths 	filePaths�  �  � k     8�� ��� l     ����  � _ Y Self-explanatory. Copy file from place A to place B. The original file will still exist.   � ��� �   S e l f - e x p l a n a t o r y .   C o p y   f i l e   f r o m   p l a c e   A   t o   p l a c e   B .   T h e   o r i g i n a l   f i l e   w i l l   s t i l l   e x i s t .� ��� r     ��� I      ���� 0 splitstring SplitString� ��� o    �� 0 	filepaths 	filePaths� ��� m    �� ���  - , -�  �  � J      �� ��� o      �� 0 
targetfile 
targetFile� ��� o      �� "0 destinationfile destinationFile�  � ��� Q    8���� k    .�� ��� I   +���
� .sysoexecTEXT���     TEXT� b    '��� b    #��� b    !��� b    ��� m    �� ���  c p� 1    �
� 
spac� l    ���� n     ��� 1     �

�
 
strq� o    �	�	 0 
targetfile 
targetFile�  �  � 1   ! "�
� 
spac� l  # &���� n   # &��� 1   $ &�
� 
strq� o   # $�� "0 destinationfile destinationFile�  �  �  � ��� L   , .�� m   , -�
� boovtrue�  � R      �� ��
� .ascrerr ****      � ****�   ��  � L   6 8�� m   6 7��
�� boovfals�  � ��� l     ��������  ��  ��  � ��� i   $ '��� I      ������� 0 createzipfile CreateZipFile� ���� o      ���� 0 paramstring paramString��  ��  � k     <�� ��� l     ������  � q k Create a ZIP file of all the PDFs in the target folder. Makes it simpler for you to send them to your KTs.   � ��� �   C r e a t e   a   Z I P   f i l e   o f   a l l   t h e   P D F s   i n   t h e   t a r g e t   f o l d e r .   M a k e s   i t   s i m p l e r   f o r   y o u   t o   s e n d   t h e m   t o   y o u r   K T s .� ��� r     ��� I      ������� 0 splitstring SplitString� ��� o    ���� 0 paramstring paramString� ���� m    �� ���  - , -��  ��  � J      �� ��� o      ���� 0 savepath savePath� ���� o      ���� 0 zippath zipPath��  � ���� Q    <���� k    2�� ��� I   /�����
�� .sysoexecTEXT���     TEXT� b    +��� b    )��� b    '��� b    #��� b    !��� b       m     �  c d 1    ��
�� 
spac� n      1     ��
�� 
strq o    ���� 0 savepath savePath� m   ! " � (   & &   / u s r / b i n / z i p   - j  � n   # &	 1   $ &��
�� 
strq	 o   # $���� 0 zippath zipPath� 1   ' (��
�� 
spac� m   ) *

 � 
 * . p d f��  � �� L   0 2 m   0 1 �  S u c c e s s��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   : < o   : ;���� 0 errmsg errMsg��  �  l     ��������  ��  ��    i   ( + I      ������ 0 
deletefile 
DeleteFile �� o      ���� 0 filepath filePath��  ��   k       l     ����   M GSelf-explanatory. This will delete the target file, skipping the Trash.    � � S e l f - e x p l a n a t o r y .   T h i s   w i l l   d e l e t e   t h e   t a r g e t   f i l e ,   s k i p p i n g   t h e   T r a s h .   l      ��!"��  ! � � The value of filePath passed to this function is always carefully considered
	(and limited), but at a future point, I will likely add in some safety checks for extra security
	to prevent a dangerous value accidentally being sent to this function.
	   " �##�   T h e   v a l u e   o f   f i l e P a t h   p a s s e d   t o   t h i s   f u n c t i o n   i s   a l w a y s   c a r e f u l l y   c o n s i d e r e d 
 	 ( a n d   l i m i t e d ) ,   b u t   a t   a   f u t u r e   p o i n t ,   I   w i l l   l i k e l y   a d d   i n   s o m e   s a f e t y   c h e c k s   f o r   e x t r a   s e c u r i t y 
 	 t o   p r e v e n t   a   d a n g e r o u s   v a l u e   a c c i d e n t a l l y   b e i n g   s e n t   t o   t h i s   f u n c t i o n . 
 	  $��$ Q     %&'% k    (( )*) I   ��+��
�� .sysoexecTEXT���     TEXT+ b    
,-, b    ./. m    00 �11 
 r m   - f/ 1    ��
�� 
spac- l   	2����2 n    	343 1    	��
�� 
strq4 o    ���� 0 filepath filePath��  ��  ��  * 5��5 L    66 m    ��
�� boovtrue��  & R      ������
�� .ascrerr ****      � ****��  ��  ' L    77 m    ��
�� boovfals��   898 l     ��������  ��  ��  9 :;: i   , /<=< I      ��>���� "0 doesbundleexist DoesBundleExist> ?��? o      ���� 0 
bundlepath 
bundlePath��  ��  = k     @@ ABA l     ��CD��  C D > Used to check if the Dialog Toolkit Plus script bundle exists   D �EE |   U s e d   t o   c h e c k   i f   t h e   D i a l o g   T o o l k i t   P l u s   s c r i p t   b u n d l e   e x i s t sB F��F O    GHG L    II l   J����J I   ��K��
�� .coredoexnull���     ****K 4    ��L
�� 
ditmL o    ���� 0 
bundlepath 
bundlePath��  ��  ��  H m     MM�                                                                                  sevs  alis    N  macOS                      �z2[BD ����System Events.app                                              �����z2[        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  ��  ; NON l     ��������  ��  ��  O PQP i   0 3RSR I      ��T���� 0 doesfileexist DoesFileExistT U��U o      ���� 0 filepath filePath��  ��  S k     VV WXW l     ��YZ��  Y   Self-explanatory   Z �[[ "   S e l f - e x p l a n a t o r yX \��\ O    ]^] L    __ F    `a` l   b����b I   ��c��
�� .coredoexnull���     ****c 4    ��d
�� 
ditmd o    ���� 0 filepath filePath��  ��  ��  a =    efe n    ghg m    ��
�� 
pclsh 4    ��i
�� 
ditmi o    ���� 0 filepath filePathf m    ��
�� 
file^ m     jj�                                                                                  sevs  alis    N  macOS                      �z2[BD ����System Events.app                                              �����z2[        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  ��  Q klk l     ��������  ��  ��  l mnm i   4 7opo I      ��q���� 0 downloadfile DownloadFileq r��r o      ���� 0 paramstring paramString��  ��  p k     Bss tut l     ��vw��  v Z T Self-explanatory. The value of fileURL is the internet address to the desired file.   w �xx �   S e l f - e x p l a n a t o r y .   T h e   v a l u e   o f   f i l e U R L   i s   t h e   i n t e r n e t   a d d r e s s   t o   t h e   d e s i r e d   f i l e .u yzy r     {|{ I      ��}���� 0 splitstring SplitString} ~~ o    ���� 0 paramstring paramString ���� m    �� ���  - , -��  ��  | J      �� ��� o      ���� "0 destinationpath destinationPath� ���� o      ���� 0 fileurl fileURL��  z ���� Q    B���� k    .�� ��� I   +�����
�� .sysoexecTEXT���     TEXT� b    '��� b    #��� b    !��� b    ��� m    �� ���  c u r l   - L   - o� 1    ��
�� 
spac� l    ������ n     ��� 1     ��
�� 
strq� o    ���� "0 destinationpath destinationPath��  ��  � 1   ! "��
�� 
spac� l  # &������ n   # &��� 1   $ &��
�� 
strq� o   # $���� 0 fileurl fileURL��  ��  ��  � ���� L   , .�� m   , -��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � k   6 B�� ��� I  6 ?�����
�� .sysodlogaskr        TEXT� b   6 ;��� b   6 9��� m   6 7�� ��� . E r r o r   d o w n l o a d i n g   f i l e :� 1   7 8��
�� 
spac� o   9 :���� 0 fileurl fileURL��  � ���� L   @ B�� m   @ A��
�� boovfals��  ��  n ��� l     ����~��  �  �~  � ��� i   8 ;��� I      �}��|�} 0 findsignature FindSignature� ��{� o      �z�z 0 signaturepath signaturePath�{  �|  � k     3�� ��� l     �y���y  � m g If your signature isn't embedded in the Excel file, it will try to find an external JPG or PNG version   � ��� �   I f   y o u r   s i g n a t u r e   i s n ' t   e m b e d d e d   i n   t h e   E x c e l   f i l e ,   i t   w i l l   t r y   t o   f i n d   a n   e x t e r n a l   J P G   o r   P N G   v e r s i o n� ��x� Q     3���� Z    )����� I    �w��v�w 0 doesfileexist DoesFileExist� ��u� b    ��� o    �t�t 0 signaturepath signaturePath� m    �� ���  m y S i g n a t u r e . p n g�u  �v  � L    �� b    ��� o    �s�s 0 signaturepath signaturePath� m    �� ���  m y S i g n a t u r e . p n g� ��� I    �r��q�r 0 doesfileexist DoesFileExist� ��p� b    ��� o    �o�o 0 signaturepath signaturePath� m    �� ���  m y S i g n a t u r e . j p g�p  �q  � ��n� L     $�� b     #��� o     !�m�m 0 signaturepath signaturePath� m   ! "�� ���  m y S i g n a t u r e . p n g�n  � L   ' )�� m   ' (�� ���  � R      �l�k�j
�l .ascrerr ****      � ****�k  �j  � L   1 3�� m   1 2�� ���  �x  � ��� l     �i�h�g�i  �h  �g  � ��� i   < ?��� I      �f��e�f 0 
renamefile 
RenameFile� ��d� o      �c�c 0 paramstring paramString�d  �e  � k     D�� ��� l     �b���b  � z t This pulls double duty for renaming a file or moving it to a new location. (It's the same process to the computer.)   � ��� �   T h i s   p u l l s   d o u b l e   d u t y   f o r   r e n a m i n g   a   f i l e   o r   m o v i n g   i t   t o   a   n e w   l o c a t i o n .   ( I t ' s   t h e   s a m e   p r o c e s s   t o   t h e   c o m p u t e r . )� ��� r     ��� I      �a��`�a 0 splitstring SplitString� ��� o    �_�_ 0 paramstring paramString� ��^� m    �� ���  - , -�^  �`  � J      �� ��� o      �]�] 0 
targetfile 
targetFile� ��\� o      �[�[ 0 newfilename newFilename�\  � ��� r       n     1    �Z
�Z 
strq n     1    �Y
�Y 
psxp o    �X�X 0 
targetfile 
targetFile o      �W�W 0 
targetfile 
targetFile�  r    &	 n    $

 1   " $�V
�V 
strq n    " 1     "�U
�U 
psxp o     �T�T 0 newfilename newFilename	 o      �S�S 0 newfilename newFilename �R Q   ' D k   * :  I  * 7�Q�P
�Q .sysoexecTEXT���     TEXT b   * 3 b   * 1 b   * / b   * - m   * + � 
 m v   - f 1   + ,�O
�O 
spac o   - .�N�N 0 
targetfile 
targetFile 1   / 0�M
�M 
spac o   1 2�L�L 0 newfilename newFilename�P    �K  L   8 :!! m   8 9�J
�J boovtrue�K   R      �I�H�G
�I .ascrerr ****      � ****�H  �G   L   B D"" m   B C�F
�F boovfals�R  � #$# l     �E�D�C�E  �D  �C  $ %&% l     �B'(�B  '   Folder Manipulation   ( �)) (   F o l d e r   M a n i p u l a t i o n& *+* l     �A�@�?�A  �@  �?  + ,-, i   @ C./. I      �>0�=�> 0 clearfolder ClearFolder0 1�<1 o      �;�; 0 foldertoempty folderToEmpty�<  �=  / k     q22 343 l     �:56�:  5 a [ Empties the target folder, but only of PDF and ZIP files. This folder will not be deleted.   6 �77 �   E m p t i e s   t h e   t a r g e t   f o l d e r ,   b u t   o n l y   o f   P D F   a n d   Z I P   f i l e s .   T h i s   f o l d e r   w i l l   n o t   b e   d e l e t e d .4 8�98 Q     q9:;9 k    g<< =>= I   �8?�7
�8 .sysoexecTEXT���     TEXT? b    @A@ b    BCB b    
DED b    FGF m    HH �II  f i n dG 1    �6
�6 
spacE l   	J�5�4J n    	KLK 1    	�3
�3 
strqL o    �2�2 0 foldertoempty folderToEmpty�5  �4  C 1   
 �1
�1 
spacA m    MM �NN : - t y p e   f   - n a m e   ' * . p d f '   - d e l e t e�7  > OPO I   "�0Q�/
�0 .sysoexecTEXT���     TEXTQ b    RSR b    TUT b    VWV b    XYX m    ZZ �[[  f i n dY 1    �.
�. 
spacW l   \�-�,\ n    ]^] 1    �+
�+ 
strq^ o    �*�* 0 foldertoempty folderToEmpty�-  �,  U 1    �)
�) 
spacS m    __ �`` : - t y p e   f   - n a m e   ' * . z i p '   - d e l e t e�/  P aba l  # #�(cd�(  c F @ It then checks for a Proofs folder and clears it of DOCX files.   d �ee �   I t   t h e n   c h e c k s   f o r   a   P r o o f s   f o l d e r   a n d   c l e a r s   i t   o f   D O C X   f i l e s .b fgf r   # (hih b   # &jkj o   # $�'�' 0 foldertoempty folderToEmptyk m   $ %ll �mm  P r o o f s /i o      �&�& 0 foldertoempty folderToEmptyg non Z   ) dpq�%�$p I   ) /�#r�"�# "0 doesfolderexist DoesFolderExistr s�!s o   * +� �  0 foldertoempty folderToEmpty�!  �"  q k   2 `tt uvu I  2 A�w�
� .sysoexecTEXT���     TEXTw b   2 =xyx b   2 ;z{z b   2 9|}| b   2 5~~ m   2 3�� ���  f i n d 1   3 4�
� 
spac} l  5 8���� n   5 8��� 1   6 8�
� 
strq� o   5 6�� 0 foldertoempty folderToEmpty�  �  { 1   9 :�
� 
spacy m   ; <�� ��� < - t y p e   f   - n a m e   ' * . d o c x '   - d e l e t e�  v ��� r   B K��� I  B I���
� .earslfdrutxt  @    file� o   B C�� 0 foldertoempty folderToEmpty� ���
� 
lfiv� m   D E�
� boovfals�  � o      ��  0 foldercontents folderContents� ��� l  L L����  � < 6 If found and empty, it then deletes the Proofs folder   � ��� l   I f   f o u n d   a n d   e m p t y ,   i t   t h e n   d e l e t e s   t h e   P r o o f s   f o l d e r� ��� Z  L `����� =  L S��� l  L Q���� I  L Q���

� .corecnte****       ****� o   L M�	�	  0 foldercontents folderContents�
  �  �  � m   Q R��  � I   V \���� 0 deletefolder DeleteFolder� ��� o   W X�� 0 foldertoempty folderToEmpty�  �  �  �  �  �%  �$  o ��� L   e g�� m   e f�
� boovtrue�  : R      �� ��
� .ascrerr ****      � ****�   ��  ; L   o q�� m   o p��
�� boovfals�9  - ��� l     ��������  ��  ��  � ��� i   D G��� I      ������� .0 clearpdfsafterzipping ClearPDFsAfterZipping� ���� o      ���� 0 foldertoempty folderToEmpty��  ��  � Q     ���� k    �� ��� I   �����
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
�� boovfals� ��� l     ��������  ��  ��  � ��� i   H K��� I      ������� 0 
copyfolder 
CopyFolder� ���� o      ���� 0 
folderpath 
folderPath��  ��  � k     8�� ��� l     ������  � o i Self-explanatory. Copy a folder (or bundle) from place A to place B. The original file will still exist.   � ��� �   S e l f - e x p l a n a t o r y .   C o p y   a   f o l d e r   ( o r   b u n d l e )   f r o m   p l a c e   A   t o   p l a c e   B .   T h e   o r i g i n a l   f i l e   w i l l   s t i l l   e x i s t .� ��� r     ��� I      ������� 0 splitstring SplitString� ��� o    ���� 0 
folderpath 
folderPath� ���� m    �� ���  - , -��  ��  � J      �� ��� o      ���� 0 targetfolder targetFolder� ���� o      ���� &0 destinationfolder destinationFolder��  � ���� Q    8���� k    .�� ��� I   +�����
�� .sysoexecTEXT���     TEXT� b    '��� b    #��� b    !��� b    ��� m    �� ���  c p   - R f� 1    ��
�� 
spac� l    ������ n     ��� 1     ��
�� 
strq� o    ���� 0 targetfolder targetFolder��  ��  � 1   ! "��
�� 
spac� l  # &������ n   # &��� 1   $ &��
�� 
strq� o   # $���� &0 destinationfolder destinationFolder��  ��  ��  � ���� L   , .�� m   , -��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   6 8�� m   6 7��
�� boovfals��  � ��� l     ��������  ��  ��  � ��� i   L O��� I      ������� 0 createfolder CreateFolder� ���� o      ���� 0 
folderpath 
folderPath��  ��  � k     ��    l     ����   \ V Self-explanatory. Needed for creating the folder for where the reports will be saved.    � �   S e l f - e x p l a n a t o r y .   N e e d e d   f o r   c r e a t i n g   t h e   f o l d e r   f o r   w h e r e   t h e   r e p o r t s   w i l l   b e   s a v e d . �� Q      k    		 

 I   ����
�� .sysoexecTEXT���     TEXT b    
 b     m     �  m k d i r   - p 1    ��
�� 
spac l   	���� n    	 1    	��
�� 
strq o    ���� 0 
folderpath 
folderPath��  ��  ��   �� L     m    ��
�� boovtrue��   R      ������
�� .ascrerr ****      � ****��  ��   L     m    ��
�� boovfals��  �  l     ��������  ��  ��    i   P S I      ������ 0 deletefolder DeleteFolder  ��  o      ���� 0 
folderpath 
folderPath��  ��   k     !! "#" l     ��$%��  $ c ] Self-explanatory. Same as with DeleteFile, extra security checks will likely be added later.   % �&& �   S e l f - e x p l a n a t o r y .   S a m e   a s   w i t h   D e l e t e F i l e ,   e x t r a   s e c u r i t y   c h e c k s   w i l l   l i k e l y   b e   a d d e d   l a t e r .# '��' Q     ()*( k    ++ ,-, I   ��.��
�� .sysoexecTEXT���     TEXT. b    
/0/ b    121 m    33 �44  r m   - r f2 1    ��
�� 
spac0 l   	5����5 n    	676 1    	��
�� 
strq7 o    ���� 0 
folderpath 
folderPath��  ��  ��  - 8��8 L    99 m    ��
�� boovtrue��  ) R      ������
�� .ascrerr ****      � ****��  ��  * L    :: m    ��
�� boovfals��   ;<; l     ��������  ��  ��  < =>= i   T W?@? I      ��A���� "0 doesfolderexist DoesFolderExistA B��B o      ���� 0 
folderpath 
folderPath��  ��  @ k     CC DED l     ��FG��  F   Self-explanatory   G �HH "   S e l f - e x p l a n a t o r yE I��I O    JKJ L    LL F    MNM l   O����O I   ��P��
�� .coredoexnull���     ****P 4    ��Q
�� 
ditmQ o    ���� 0 
folderpath 
folderPath��  ��  ��  N =    RSR n    TUT m    ��
�� 
pclsU 4    ��V
�� 
ditmV o    ���� 0 
folderpath 
folderPathS m    ��
�� 
cfolK m     WW�                                                                                  sevs  alis    N  macOS                      �z2[BD ����System Events.app                                              �����z2[        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  ��  > XYX l     ��������  ��  ��  Y Z[Z l     ��\]��  \   Dialog Boxes   ] �^^    D i a l o g   B o x e s[ _`_ l     ��������  ��  ��  ` aba i   X [cdc I      ��e��� 80 installdialogdisplayscript InstallDialogDisplayScripte f�~f o      �}�} 0 paramstring paramString�~  �  d k     gg hih r     jkj b     	lml n     non 1    �|
�| 
psxpo l    p�{�zp I    �yq�x
�y .earsffdralis        afdrq m     �w
�w afdrcusr�x  �{  �z  m m    rr �ss � L i b r a r y / A p p l i c a t i o n   S c r i p t s / c o m . m i c r o s o f t . E x c e l / D i a l o g D i s p l a y . s c p tk o      �v�v 0 
scriptpath 
scriptPathi tut r    vwv m    xx �yy � h t t p s : / / r a w . g i t h u b u s e r c o n t e n t . c o m / p a p e r c u t t e r 0 3 2 4 / S p e a k i n g E v a l s / m a i n / D i a l o g D i s p l a y . s c p tw o      �u�u 0 downloadurl downloadURLu z{z l   �t�s�r�t  �s  �r  { |}| l   �q~�q  ~ A ; If an existing version is not found, download a fresh copy    ��� v   I f   a n   e x i s t i n g   v e r s i o n   i s   n o t   f o u n d ,   d o w n l o a d   a   f r e s h   c o p y} ��� l   �p���p  � e _ Skip this first check until a full update function can be designed. For now, install each time   � ��� �   S k i p   t h i s   f i r s t   c h e c k   u n t i l   a   f u l l   u p d a t e   f u n c t i o n   c a n   b e   d e s i g n e d .   F o r   n o w ,   i n s t a l l   e a c h   t i m e� ��� l   �o���o  � 4 . if DoesFileExist(scriptPath) then return true   � ��� \   i f   D o e s F i l e E x i s t ( s c r i p t P a t h )   t h e n   r e t u r n   t r u e� ��n� L    �� I    �m��l�m 0 downloadfile DownloadFile� ��k� b    ��� b    ��� o    �j�j 0 
scriptpath 
scriptPath� m    �� ���  - , -� o    �i�i 0 downloadurl downloadURL�k  �l  �n  b ��� l     �h�g�f�h  �g  �f  � ��� i   \ _��� I      �e��d�e >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder� ��c� o      �b�b 0 paramstring paramString�c  �d  � k     ]�� ��� r     ��� b     	��� n     ��� 1    �a
�a 
psxp� l    ��`�_� I    �^��]
�^ .earsffdralis        afdr� m     �\
�\ afdrcusr�]  �`  �_  � m    �� ��� 0 L i b r a r y / S c r i p t   L i b r a r i e s� o      �[�[ .0 scriptlibrariesfolder scriptLibrariesFolder� ��� l   �Z�Y�X�Z  �Y  �X  � ��W� Z    ]���V�� I    �U��T�U "0 doesfolderexist DoesFolderExist� ��S� o    �R�R .0 scriptlibrariesfolder scriptLibrariesFolder�S  �T  � L    �� o    �Q�Q .0 scriptlibrariesfolder scriptLibrariesFolder�V  � Q    ]���� k    Q�� ��� l   �P���P  � m g ~/Library is typically a read-only folder, so I need to requst your password to create the need folder   � ��� �   ~ / L i b r a r y   i s   t y p i c a l l y   a   r e a d - o n l y   f o l d e r ,   s o   I   n e e d   t o   r e q u s t   y o u r   p a s s w o r d   t o   c r e a t e   t h e   n e e d   f o l d e r� ��� I   *�O��
�O .sysoexecTEXT���     TEXT� b    $��� b     ��� m    �� ���  m k d i r   - p� 1    �N
�N 
spac� n     #��� 1   ! #�M
�M 
strq� o     !�L�L .0 scriptlibrariesfolder scriptLibrariesFolder� �K��J
�K 
badm� m   % &�I
�I boovtrue�J  � ��� l  + +�H���H  � %  Set your username as the owner   � ��� >   S e t   y o u r   u s e r n a m e   a s   t h e   o w n e r� ��� I  + B�G��
�G .sysoexecTEXT���     TEXT� b   + <��� b   + 8��� b   + 6��� m   + ,�� ���  c h o w n  � n   , 5��� 1   3 5�F
�F 
strq� l  , 3��E�D� n   , 3��� 1   1 3�C
�C 
sisn� l  , 1��B�A� I  , 1�@�?�>
�@ .sysosigtsirr   ��� null�?  �>  �B  �A  �E  �D  � 1   6 7�=
�= 
spac� n   8 ;��� 1   9 ;�<
�< 
strq� o   8 9�;�; .0 scriptlibrariesfolder scriptLibrariesFolder� �:��9
�: 
badm� m   = >�8
�8 boovtrue�9  � ��� l  C C�7���7  � 5 / Give your username READ and WRITE permissions.   � ��� ^   G i v e   y o u r   u s e r n a m e   R E A D   a n d   W R I T E   p e r m i s s i o n s .� ��� I  C N�6��
�6 .sysoexecTEXT���     TEXT� b   C H��� m   C D�� ���  c h m o d   u + r w  � n   D G��� 1   E G�5
�5 
strq� o   D E�4�4 .0 scriptlibrariesfolder scriptLibrariesFolder� �3��2
�3 
badm� m   I J�1
�1 boovtrue�2  � ��0� L   O Q�� o   O P�/�/ .0 scriptlibrariesfolder scriptLibrariesFolder�0  � R      �.�-�,
�. .ascrerr ****      � ****�-  �,  � L   Y ]�� m   Y \�� ���  �W  � ��� l     �+�*�)�+  �*  �)  � ��� i   ` c��� I      �(��'�( 40 installdialogtoolkitplus InstallDialogToolkitPlus� ��&� o      �%�% "0 resourcesfolder resourcesFolder�&  �'  � k     �    r      m      � � h t t p s : / / r a w . g i t h u b u s e r c o n t e n t . c o m / p a p e r c u t t e r 0 3 2 4 / S p e a k i n g E v a l s / m a i n / D i a l o g _ T o o l k i t . z i p o      �$�$ 0 downloadurl downloadURL  r    	
	 b     n     1   	 �#
�# 
psxp l   	�"�! I   	� �
�  .earsffdralis        afdr m    �
� afdrcusr�  �"  �!   m     � 0 L i b r a r y / S c r i p t   L i b r a r i e s
 o      �� .0 scriptlibrariesfolder scriptLibrariesFolder  r     m     � 4 / D i a l o g   T o o l k i t   P l u s . s c p t d o      �� $0 dialogbundlename dialogBundleName  r     b     o    �� .0 scriptlibrariesfolder scriptLibrariesFolder o    �� $0 dialogbundlename dialogBundleName o      �� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle   r    !"! b    #$# o    �� "0 resourcesfolder resourcesFolder$ m    %% �&& & / D i a l o g _ T o o l k i t . z i p" o      �� 0 zipfilepath zipFilePath  '(' r     %)*) b     #+,+ o     !�� "0 resourcesfolder resourcesFolder, m   ! "-- �.. $ / d i a l o g T o o l k i t T e m p* o      �� &0 zipextractionpath zipExtractionPath( /0/ l  & &����  �  �  0 121 l  & &�34�  3 0 * Initial check to see if already installed   4 �55 T   I n i t i a l   c h e c k   t o   s e e   i f   a l r e a d y   i n s t a l l e d2 676 Z  & 589��8 I   & ,�:�� "0 doesbundleexist DoesBundleExist: ;�; o   ' (�� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�  �  9 L   / 1<< m   / 0�

�
 boovtrue�  �  7 =>= l  6 6�	���	  �  �  > ?@? l  6 6�AB�  A 3 - Ensure resources folder exists for later use   B �CC Z   E n s u r e   r e s o u r c e s   f o l d e r   e x i s t s   f o r   l a t e r   u s e@ DED Z   6 WFG��F H   6 =HH I   6 <�I�� "0 doesfolderexist DoesFolderExistI J�J o   7 8� �  "0 resourcesfolder resourcesFolder�  �  G Q   @ SKLMK I   C I��N���� 0 createfolder CreateFolderN O��O o   D E���� "0 resourcesfolder resourcesFolder��  ��  L R      ������
�� .ascrerr ****      � ****��  ��  M L   Q SPP m   Q R��
�� boovfals�  �  E QRQ l  X X��������  ��  ��  R STS l  X X��UV��  U G A Check for a local copy and move it to the needed folder if found   V �WW �   C h e c k   f o r   a   l o c a l   c o p y   a n d   m o v e   i t   t o   t h e   n e e d e d   f o l d e r   i f   f o u n dT XYX Z   X |Z[����Z I   X `��\���� "0 doesbundleexist DoesBundleExist\ ]��] b   Y \^_^ o   Y Z���� "0 resourcesfolder resourcesFolder_ o   Z [���� $0 dialogbundlename dialogBundleName��  ��  [ Z   c x`a����` I   c o��b���� 0 
copyfolder 
CopyFolderb c��c b   d kded b   d ifgf b   d ghih o   d e���� "0 resourcesfolder resourcesFolderi o   e f���� $0 dialogbundlename dialogBundleNameg m   g hjj �kk  - , -e o   i j���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  a L   r tll m   r s��
�� boovtrue��  ��  ��  ��  Y mnm l  } }��������  ��  ��  n opo l  } }��qr��  q !  Otherwise, download and...   r �ss 6   O t h e r w i s e ,   d o w n l o a d   a n d . . .p tut Z   } �vw����v I   } ���x���� 0 downloadfile DownloadFilex y��y b   ~ �z{z b   ~ �|}| o   ~ ���� 0 zipfilepath zipFilePath} m    �~~ �  - , -{ o   � ����� 0 downloadurl downloadURL��  ��  w Q   � ������ k   � ��� ��� l  � �������  �   ...extract the files...   � ��� 0   . . . e x t r a c t   t h e   f i l e s . . .� ��� I  � ������
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
�� .ascrerr ****      � ****��  ��  ��  ��  ��  u ��� l  � ���������  ��  ��  � ��� l  � �������  � D > Remove unneeded files and folders created during this process   � ��� |   R e m o v e   u n n e e d e d   f i l e s   a n d   f o l d e r s   c r e a t e d   d u r i n g   t h i s   p r o c e s s� ��� I   � �������� 0 
deletefile 
DeleteFile� ���� o   � ����� 0 zipfilepath zipFilePath��  ��  � ��� I   � �������� 0 deletefolder DeleteFolder� ���� o   � ����� &0 zipextractionpath zipExtractionPath��  ��  � ��� l  � ���������  ��  ��  � ��� l  � �������  � V P One final check to verify installation was successful and return true if it was   � ��� �   O n e   f i n a l   c h e c k   t o   v e r i f y   i n s t a l l a t i o n   w a s   s u c c e s s f u l   a n d   r e t u r n   t r u e   i f   i t   w a s� ���� L   � ��� I   � �������� "0 doesbundleexist DoesBundleExist� ���� o   � ����� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  ��  � ��� l     ��������  ��  ��  � ��� i   d g��� I      ������� 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus� ���� o      ���� "0 resourcesfolder resourcesFolder��  ��  � k     U�� ��� r     ��� b     	��� n     ��� 1    ��
�� 
psxp� l    ������ I    �����
�� .earsffdralis        afdr� m     ��
�� afdrcusr��  ��  ��  � m    �� ��� d L i b r a r y / S c r i p t   L i b r a r i e s / D i a l o g   T o o l k i t   P l u s . s c p t d� o      ���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle� ��� r    ��� b    ��� o    ���� "0 resourcesfolder resourcesFolder� m    �� ��� 4 / D i a l o g   T o o l k i t   P l u s . s c p t d� o      ���� 0 	localcopy 	localCopy� � � l   ��������  ��  ��     Z    R�� I    ������ "0 doesbundleexist DoesBundleExist �� o    ���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��   Q    L	
 k    A  Z   6���� H    % I    $������ "0 doesbundleexist DoesBundleExist �� o     ���� 0 	localcopy 	localCopy��  ��   I   ( 2������ 0 
copyfolder 
CopyFolder �� b   ) . b   ) , o   ) *���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle m   * + �  - , - o   , -���� 0 	localcopy 	localCopy��  ��  ��  ��    I   7 =����� 0 deletefolder DeleteFolder �~ o   8 9�}�} 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�~  �   �| r   > A !  m   > ?�{
�{ boovtrue! o      �z�z 0 removalresult removalResult�|  	 R      �y�x�w
�y .ascrerr ****      � ****�x  �w  
 r   I L"#" m   I J�v
�v boovfals# o      �u�u 0 removalresult removalResult��   r   O R$%$ m   O P�t
�t boovtrue% o      �s�s 0 removalresult removalResult &'& l  S S�r�q�p�r  �q  �p  ' (�o( L   S U)) o   S T�n�n 0 removalresult removalResult�o  � *�m* l     �l�k�j�l  �k  �j  �m       �i+,-./0123456789:;<=>?@ABCDE�i  + �h�g�f�e�d�c�b�a�`�_�^�]�\�[�Z�Y�X�W�V�U�T�S�R�Q�P�O�h 00 getscriptversionnumber GetScriptVersionNumber�g "0 getmacosversion GetMacOSVersion�f 80 checkaccessibilitysettings CheckAccessibilitySettings�e 0 splitstring SplitString�d "0 loadapplication LoadApplication�c 0 isapploaded IsAppLoaded�b 0 	closeword 	CloseWord�a $0 comparemd5hashes CompareMD5Hashes�` 0 copyfile CopyFile�_ 0 createzipfile CreateZipFile�^ 0 
deletefile 
DeleteFile�] "0 doesbundleexist DoesBundleExist�\ 0 doesfileexist DoesFileExist�[ 0 downloadfile DownloadFile�Z 0 findsignature FindSignature�Y 0 
renamefile 
RenameFile�X 0 clearfolder ClearFolder�W .0 clearpdfsafterzipping ClearPDFsAfterZipping�V 0 
copyfolder 
CopyFolder�U 0 createfolder CreateFolder�T 0 deletefolder DeleteFolder�S "0 doesfolderexist DoesFolderExist�R 80 installdialogdisplayscript InstallDialogDisplayScript�Q >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder�P 40 installdialogtoolkitplus InstallDialogToolkitPlus�O 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus, �N �M�LFG�K�N 00 getscriptversionnumber GetScriptVersionNumber�M �JH�J H  �I�I 0 paramstring paramString�L  F �H�H 0 paramstring paramStringG �G�G 4�_�K �- �F &�E�DIJ�C�F "0 getmacosversion GetMacOSVersion�E �BK�B K  �A�A 0 paramstring paramString�D  I �@�?�@ 0 paramstring paramString�? 0 	osversion 	osVersionJ  8�>�=�<
�> .sysoexecTEXT���     TEXT�=  �<  �C  �j E�O�W X  h. �; A�:�9LM�8�; 80 checkaccessibilitysettings CheckAccessibilitySettings�: �7N�7 N  �6�6 0 
apptocheck 
appToCheck�9  L �5�4�5 0 
apptocheck 
appToCheck�4 ,0 accessibilityenabled accessibilityEnabledM  n�3�2O�1�0�/�.�-�,�+
�3 
prcs
�2 
pnamO  
�1 
pvis
�0 
pcap
�/ 
uiel
�. 
enaB
�- 
bool�,  �+  �8 6 -� %*�-�,�[�,\Ze81�	 *�/�-�,E�&E�O�UW 	X 	 
f/ �* |�)�(PQ�'�* 0 splitstring SplitString�) �&R�& R  �%�$�% &0 passedparamstring passedParamString�$ (0 parameterseparator parameterSeparator�(  P �#�"�!� �# &0 passedparamstring passedParamString�" (0 parameterseparator parameterSeparator�! 00 oldtextitemsdelimiters oldTextItemsDelimiters�  *0 separatedparameters separatedParametersQ ���
� 
ascr
� 
txdl
� 
citm�'  � *�,E�O�*�,FO��-E�O�*�,FUO�0 � ���ST�� "0 loadapplication LoadApplication� �U� U  �� 0 appname appName�  S ���� 0 appname appName� 0 errmsg errMsg� 0 errnum errNumT 	�� ��V �� � �
� 
capp
� .miscactvnull��� ��� null� 0 errmsg errMsgV ���
� 
errn� 0 errnum errNum�  
� 
spac� * *�/ *j UO�W X  ��%�%�%�%�%�%1 � ���
WX�	� 0 isapploaded IsAppLoaded� �Y� Y  �� 0 appname appName�
  W ����� 0 appname appName� 0 
loadresult 
loadResult� 0 errmsg errMsg� 0 errnum errNumX ���  ���Z
� 
prcs
� 
pnam
�  
spac�� 0 errmsg errMsgZ ������
�� 
errn�� 0 errnum errNum��  �	 ; (� *�-�,� ��%�%E�Y 	��%�%E�UO�W X  �%�%�%�%�%2 ��#����[\���� 0 	closeword 	CloseWord�� ��]�� ]  ���� 0 paramstring paramString��  [ ������ 0 paramstring paramString�� 0 closeresult closeResult\ P����=D��HL����R
�� 
prcs
�� 
pnam
�� .aevtquitnull��� ��� null��  ��  �� 4 +� #*�-�,� � *j UO�E�Y �E�O�UW 	X  	�3 ��`����^_���� $0 comparemd5hashes CompareMD5Hashes�� ��`�� `  ���� 0 paramstring paramString��  ^ ���������� 0 paramstring paramString�� 0 filepath filePath�� 0 	validhash 	validHash�� 0 checkresult checkResult_ 
q������������������� 0 splitstring SplitString
�� 
cobj�� 0 doesfileexist DoesFileExist
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� H*��l+ E[�k/E�Z[�l/E�ZO*�k+  fY hO ��%��,%j E�O�� W 	X  	f4 �������ab���� 0 copyfile CopyFile�� ��c�� c  ���� 0 	filepaths 	filePaths��  a �������� 0 	filepaths 	filePaths�� 0 
targetfile 
targetFile�� "0 destinationfile destinationFileb 	������������������ 0 splitstring SplitString
�� 
cobj
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� 9*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW 	X  f5 �������de���� 0 createzipfile CreateZipFile�� ��f�� f  ���� 0 paramstring paramString��  d ���������� 0 paramstring paramString�� 0 savepath savePath�� 0 zippath zipPath�� 0 errmsg errMsge ���������
�������� 0 splitstring SplitString
�� 
cobj
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� =*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%�%�%j O�W 	X 
 �6 ������gh���� 0 
deletefile 
DeleteFile�� ��i�� i  ���� 0 filepath filePath��  g ���� 0 filepath filePathh 0����������
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  ��  ��%��,%j OeW 	X  f7 ��=����jk���� "0 doesbundleexist DoesBundleExist�� ��l�� l  ���� 0 
bundlepath 
bundlePath��  j ���� 0 
bundlepath 
bundlePathk M����
�� 
ditm
�� .coredoexnull���     ****�� � *�/j U8 ��S����mn���� 0 doesfileexist DoesFileExist�� ��o�� o  ���� 0 filepath filePath��  m ���� 0 filepath filePathn j����������
�� 
ditm
�� .coredoexnull���     ****
�� 
pcls
�� 
file
�� 
bool�� � *�/j 	 *�/�,� �&U9 ��p����pq���� 0 downloadfile DownloadFile�� ��r�� r  ���� 0 paramstring paramString��  p �������� 0 paramstring paramString�� "0 destinationpath destinationPath�� 0 fileurl fileURLq ��������������������� 0 splitstring SplitString
�� 
cobj
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  
�� .sysodlogaskr        TEXT�� C*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW X  ��%�%j 
Of: �������st���� 0 findsignature FindSignature�� ��u�� u  ���� 0 signaturepath signaturePath��  s ���� 0 signaturepath signaturePatht 	�������������� 0 doesfileexist DoesFileExist��  ��  �� 4 +*��%k+  	��%Y *��%k+  	��%Y �W 	X  �; ���~�}vw�|� 0 
renamefile 
RenameFile�~ �{x�{ x  �z�z 0 paramstring paramString�}  v �y�x�w�y 0 paramstring paramString�x 0 
targetfile 
targetFile�w 0 newfilename newFilenamew 
��v�u�t�s�r�q�p�o�v 0 splitstring SplitString
�u 
cobj
�t 
psxp
�s 
strq
�r 
spac
�q .sysoexecTEXT���     TEXT�p  �o  �| E*��l+ E[�k/E�Z[�l/E�ZO��,�,E�O��,�,E�O ��%�%�%�%j OeW 	X  	f< �n/�m�lyz�k�n 0 clearfolder ClearFolder�m �j{�j {  �i�i 0 foldertoempty folderToEmpty�l  y �h�g�h 0 foldertoempty folderToEmpty�g  0 foldercontents folderContentsz H�f�eM�dZ_l�c���b�a�`�_�^�]
�f 
spac
�e 
strq
�d .sysoexecTEXT���     TEXT�c "0 doesfolderexist DoesFolderExist
�b 
lfiv
�a .earslfdrutxt  @    file
�` .corecnte****       ****�_ 0 deletefolder DeleteFolder�^  �]  �k r i��%��,%�%�%j O��%��,%�%�%j O��%E�O*�k+  3��%��,%�%�%j O��fl E�O�j j  *�k+ Y hY hOeW 	X  f= �\��[�Z|}�Y�\ .0 clearpdfsafterzipping ClearPDFsAfterZipping�[ �X~�X ~  �W�W 0 foldertoempty folderToEmpty�Z  | �V�V 0 foldertoempty folderToEmpty} ��U�T��S�R�Q
�U 
spac
�T 
strq
�S .sysoexecTEXT���     TEXT�R  �Q  �Y   ��%��,%�%�%j OeW 	X  f> �P��O�N��M�P 0 
copyfolder 
CopyFolder�O �L��L �  �K�K 0 
folderpath 
folderPath�N   �J�I�H�J 0 
folderpath 
folderPath�I 0 targetfolder targetFolder�H &0 destinationfolder destinationFolder� 	��G�F��E�D�C�B�A�G 0 splitstring SplitString
�F 
cobj
�E 
spac
�D 
strq
�C .sysoexecTEXT���     TEXT�B  �A  �M 9*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW 	X  f? �@��?�>���=�@ 0 createfolder CreateFolder�? �<��< �  �;�; 0 
folderpath 
folderPath�>  � �:�: 0 
folderpath 
folderPath� �9�8�7�6�5
�9 
spac
�8 
strq
�7 .sysoexecTEXT���     TEXT�6  �5  �=  ��%��,%j OeW 	X  f@ �4�3�2���1�4 0 deletefolder DeleteFolder�3 �0��0 �  �/�/ 0 
folderpath 
folderPath�2  � �.�. 0 
folderpath 
folderPath� 3�-�,�+�*�)
�- 
spac
�, 
strq
�+ .sysoexecTEXT���     TEXT�*  �)  �1  ��%��,%j OeW 	X  fA �(@�'�&���%�( "0 doesfolderexist DoesFolderExist�' �$��$ �  �#�# 0 
folderpath 
folderPath�&  � �"�" 0 
folderpath 
folderPath� W�!� ���
�! 
ditm
�  .coredoexnull���     ****
� 
pcls
� 
cfol
� 
bool�% � *�/j 	 *�/�,� �&UB �d������ 80 installdialogdisplayscript InstallDialogDisplayScript� ��� �  �� 0 paramstring paramString�  � ���� 0 paramstring paramString� 0 
scriptpath 
scriptPath� 0 downloadurl downloadURL� ���rx��
� afdrcusr
� .earsffdralis        afdr
� 
psxp� 0 downloadfile DownloadFile� �j �,�%E�O�E�O*��%�%k+ C �������� >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder� ��� �  �
�
 0 paramstring paramString�  � �	��	 0 paramstring paramString� .0 scriptlibrariesfolder scriptLibrariesFolder� ���������� �����������
� afdrcusr
� .earsffdralis        afdr
� 
psxp� "0 doesfolderexist DoesFolderExist
� 
spac
� 
strq
� 
badm
�  .sysoexecTEXT���     TEXT
�� .sysosigtsirr   ��� null
�� 
sisn��  ��  � ^�j �,�%E�O*�k+  �Y E 9��%��,%�el 	O�*j �,�,%�%��,%�el 	O���,%�el 	O�W X  a D ������������� 40 installdialogtoolkitplus InstallDialogToolkitPlus�� ����� �  ���� "0 resourcesfolder resourcesFolder��  � ���������������� "0 resourcesfolder resourcesFolder�� 0 downloadurl downloadURL�� .0 scriptlibrariesfolder scriptLibrariesFolder�� $0 dialogbundlename dialogBundleName�� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�� 0 zipfilepath zipFilePath�� &0 zipextractionpath zipExtractionPath� ������%-����������j��~������������������
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
W 	X  fY hO*��%k+  *��%�%�%k+  eY hY hO*��%�%k+  T Ha _ %�a ,%a %�a ,%j O*�a %�%a %�%�%k+ O*�a %�%a %�%k+ W X  hY hO*�k+ O*�k+ O*�k+ E ������������� 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus�� ����� �  ���� "0 resourcesfolder resourcesFolder��  � ���������� "0 resourcesfolder resourcesFolder�� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�� 0 	localcopy 	localCopy�� 0 removalresult removalResult� ������������������
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp�� "0 doesbundleexist DoesBundleExist�� 0 
copyfolder 
CopyFolder�� 0 deletefolder DeleteFolder��  ��  �� V�j �,�%E�O��%E�O*�k+  6 (*�k+  *��%�%k+ Y hO*�k+ OeE�W 
X 	 
fE�Y eE�O� ascr  ��ޭ