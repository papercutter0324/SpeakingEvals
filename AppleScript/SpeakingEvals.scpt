FasdUAS 1.101.10   ��   ��    k             l      ��  ��    � |
Helper Scripts for the DYB Speaking Evaluations Excel spreadsheet

Version: 1.5.0
Build:   20250704
Warren Feltmate
� 2025
     � 	 	 � 
 H e l p e r   S c r i p t s   f o r   t h e   D Y B   S p e a k i n g   E v a l u a t i o n s   E x c e l   s p r e a d s h e e t 
 
 V e r s i o n :   1 . 5 . 0 
 B u i l d :       2 0 2 5 0 7 0 4 
 W a r r e n   F e l t m a t e 
 �   2 0 2 5 
   
  
 l     ��������  ��  ��        l     ��  ��      Environment Variables     �   ,   E n v i r o n m e n t   V a r i a b l e s      l     ��������  ��  ��        i         I      �� ���� 00 getscriptversionnumber GetScriptVersionNumber   ��  o      ���� 0 paramstring paramString��  ��    k            l     ��  ��    ? 9- Use build number to determine if an update is available     �   r -   U s e   b u i l d   n u m b e r   t o   d e t e r m i n e   i f   a n   u p d a t e   i s   a v a i l a b l e   ��  L          m     ���� 5 P��     ! " ! l     ��������  ��  ��   "  # $ # i     % & % I      �� '���� "0 getmacosversion GetMacOSVersion '  (�� ( o      ���� 0 paramstring paramString��  ��   & k      ) )  * + * l     �� , -��   , ` Z Not currently used, but could be helpful if there are issues with older versions of MacOS    - � . . �   N o t   c u r r e n t l y   u s e d ,   b u t   c o u l d   b e   h e l p f u l   i f   t h e r e   a r e   i s s u e s   w i t h   o l d e r   v e r s i o n s   o f   M a c O S +  /�� / Q      0 1�� 0 k     2 2  3 4 3 r    
 5 6 5 I   �� 7��
�� .sysoexecTEXT���     TEXT 7 m     8 8 � 9 9 . s w _ v e r s   - p r o d u c t V e r s i o n��   6 o      ���� 0 	osversion 	osVersion 4  :�� : L     ; ; o    ���� 0 	osversion 	osVersion��   1 R      ������
�� .ascrerr ****      � ****��  ��  ��  ��   $  < = < l     ��������  ��  ��   =  > ? > l     �� @ A��   @   Parameter Manipulation    A � B B .   P a r a m e t e r   M a n i p u l a t i o n ?  C D C l     ��������  ��  ��   D  E F E i     G H G I      �� I���� 0 splitstring SplitString I  J K J o      ���� &0 passedparamstring passedParamString K  L�� L o      ���� (0 parameterseparator parameterSeparator��  ��   H k      M M  N O N l     �� P Q��   P d ^ Excel can only pass on parameter to this file. This makes it possible to split one into many.    Q � R R �   E x c e l   c a n   o n l y   p a s s   o n   p a r a m e t e r   t o   t h i s   f i l e .   T h i s   m a k e s   i t   p o s s i b l e   t o   s p l i t   o n e   i n t o   m a n y . O  S T S O      U V U k     W W  X Y X r    	 Z [ Z n    \ ] \ 1    ��
�� 
txdl ] 1    ��
�� 
ascr [ o      ���� 00 oldtextitemsdelimiters oldTextItemsDelimiters Y  ^ _ ^ r   
  ` a ` o   
 ���� (0 parameterseparator parameterSeparator a n      b c b 1    ��
�� 
txdl c 1    ��
�� 
ascr _  d e d r     f g f n     h i h 2   ��
�� 
citm i o    ���� &0 passedparamstring passedParamString g o      ���� *0 separatedparameters separatedParameters e  j�� j r     k l k o    ���� 00 oldtextitemsdelimiters oldTextItemsDelimiters l n      m n m 1    ��
�� 
txdl n 1    ��
�� 
ascr��   V 1     ��
�� 
ascr T  o�� o L     p p o    ���� *0 separatedparameters separatedParameters��   F  q r q l     ��������  ��  ��   r  s t s i     u v u I      �� w���� 0 
joinstring 
JoinString w  x y x o      ���� $0 passedparamarray passedParamArray y  z�� z o      ���� (0 parameterseparator parameterSeparator��  ��   v k      { {  | } | O      ~  ~ k     � �  � � � r    	 � � � n    � � � 1    ��
�� 
txdl � 1    ��
�� 
ascr � o      ���� 00 oldtextitemsdelimiters oldTextItemsDelimiters �  � � � r   
  � � � o   
 ���� (0 parameterseparator parameterSeparator � n      � � � 1    ��
�� 
txdl � 1    ��
�� 
ascr �  � � � r     � � � c     � � � o    ���� $0 passedparamarray passedParamArray � m    ��
�� 
TEXT � o      ���� $0 joinedparameters joinedParameters �  ��� � r     � � � o    ���� 00 oldtextitemsdelimiters oldTextItemsDelimiters � n      � � � 1    ��
�� 
txdl � 1    ��
�� 
ascr��    1     ��
�� 
ascr }  ��� � L     � � o    ���� $0 joinedparameters joinedParameters��   t  � � � l     ��������  ��  ��   �  � � � l     �� � ���   �    Application Manipulations    � � � � 4   A p p l i c a t i o n   M a n i p u l a t i o n s �  � � � l     ��������  ��  ��   �  � � � i     � � � I      �� ����� "0 loadapplication LoadApplication �  ��� � o      ���� 0 appname appName��  ��   � k     ) � �  � � � l     �� � ���   � < 6 A simple function to tell the needed program to open.    � � � � l   A   s i m p l e   f u n c t i o n   t o   t e l l   t h e   n e e d e d   p r o g r a m   t o   o p e n . �  ��� � Q     ) � � � � k     � �  � � � O    � � � I  
 ������
�� .miscactvnull��� ��� null��  ��   � 4    �� �
�� 
capp � o    ���� 0 appname appName �  ��� � L     � � m     � � � � �  ��   � R      �� � �
�� .ascrerr ****      � **** � o      ���� 0 errmsg errMsg � �� ���
�� 
errn � o      ���� 0 errnum errNum��   � L    ) � � b    ( � � � b    & � � � b    $ � � � b    " � � � b      � � � b     � � � m     � � � � �  E r r o r   l o a d i n g � 1    ��
�� 
spac � o    ���� 0 appname appName � m     ! � � � � �  :   � o   " #���� 0 errnum errNum � m   $ % � � � � �    -   � o   & '���� 0 errmsg errMsg��   �  � � � l     ��������  ��  ��   �  � � � i     � � � I      �� ����� 0 isapploaded IsAppLoaded �  ��� � o      ���� 0 appname appName��  ��   � k     : � �  � � � l     �� � ���   � N H This lets Excel check that the other program is open before continuing.    � � � � �   T h i s   l e t s   E x c e l   c h e c k   t h a t   t h e   o t h e r   p r o g r a m   i s   o p e n   b e f o r e   c o n t i n u i n g . �  ��� � Q     : � � � � k    & � �  � � � O    # � � � Z    " � ��� � � E     � � � l    ����� � n     � � � 1   
 ��
�� 
pnam � 2    
�
� 
prcs��  ��   � o    �~�~ 0 appname appName � r     � � � b     � � � b     � � � o    �}�} 0 appname appName � 1    �|
�| 
spac � m     � � � � �  i s   n o w   r u n n i n g . � o      �{�{ 0 
loadresult 
loadResult��   � r    " � � � b      � � � b     � � � m       �  E r r o r   o p e n i n g � 1    �z
�z 
spac � o    �y�y 0 appname appName � o      �x�x 0 
loadresult 
loadResult � m    �                                                                                  sevs  alis    \  Macintosh HD               �=,�BD ����System Events.app                                              �����=,�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��   � �w L   $ & o   $ %�v�v 0 
loadresult 
loadResult�w   � R      �u
�u .ascrerr ****      � **** o      �t�t 0 errmsg errMsg �s�r
�s 
errn o      �q�q 0 errnum errNum�r   � L   . : b   . 9	
	 b   . 7 b   . 5 b   . 3 b   . 1 m   . / �  E r r o r   l o a d i n g   o   / 0�p�p 0 appname appName m   1 2 �  :   o   3 4�o�o 0 errnum errNum m   5 6 �    -  
 o   7 8�n�n 0 errmsg errMsg��   �  l     �m�l�k�m  �l  �k    i     I      �j�i�j "0 closepowerpoint ClosePowerPoint  �h  o      �g�g 0 paramstring paramString�h  �i   k     3!! "#" l     �f$%�f  $ { u This will completely close MS PowerPoint, even from the Dock. This reduces the chances of errors on subsequent runs.   % �&& �   T h i s   w i l l   c o m p l e t e l y   c l o s e   M S   P o w e r P o i n t ,   e v e n   f r o m   t h e   D o c k .   T h i s   r e d u c e s   t h e   c h a n c e s   o f   e r r o r s   o n   s u b s e q u e n t   r u n s .# '�e' Q     3()*( O    )+,+ k    (-- ./. Z    %01�d20 E    343 l   5�c�b5 n    676 1   
 �a
�a 
pnam7 2    
�`
�` 
prcs�c  �b  4 m    88 �99 ( M i c r o s o f t   P o w e r P o i n t1 k    :: ;<; O   =>= I   �_�^�]
�_ .aevtquitnull��� ��� null�^  �]  > m    ??�                                                                                  PPT3  alis    Z  Macintosh HD               �=,�BD ����Microsoft PowerPoint.app                                       �����kd         ����  
 cu             Applications  (/:Applications:Microsoft PowerPoint.app/  2  M i c r o s o f t   P o w e r P o i n t . a p p    M a c i n t o s h   H D  %Applications/Microsoft PowerPoint.app   / ��  < @�\@ r    ABA m    CC �DD P P o w e r P o i n t   h a s   s u c c e s s f u l l y   b e e n   c l o s e d .B o      �[�[ 0 closeresult closeResult�\  �d  2 r   " %EFE m   " #GG �HH H P o w e r P o i n t   i s   n o t   c u r r e n t l y   r u n n i n g .F o      �Z�Z 0 closeresult closeResult/ I�YI L   & (JJ o   & '�X�X 0 closeresult closeResult�Y  , m    KK�                                                                                  sevs  alis    \  Macintosh HD               �=,�BD ����System Events.app                                              �����=,�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  ) R      �W�V�U
�W .ascrerr ****      � ****�V  �U  * L   1 3LL m   1 2MM �NN \ T h e r e   w a s   a n   e r r o r   t r y i n g   t o   c l o s e   P o w e r P o i n t .�e   OPO l     �T�S�R�T  �S  �R  P QRQ l     �QST�Q  S   File Manipulation   T �UU $   F i l e   M a n i p u l a t i o nR VWV l     �P�O�N�P  �O  �N  W XYX i    Z[Z I      �M\�L�M .0 changefilepermissions ChangeFilePermissions\ ]�K] o      �J�J 0 paramstring paramString�K  �L  [ k     f^^ _`_ r     aba I      �Ic�H�I 0 splitstring SplitStringc ded o    �G�G 0 paramstring paramStringe f�Ff m    gg �hh  - , -�F  �H  b J      ii jkj o      �E�E  0 newpermissions newPermissionsk l�Dl o      �C�C 0 filepath filePath�D  ` mnm l   �B�A�@�B  �A  �@  n opo l   �?qr�?  q = 7 Check if quarantine status is set; remove if necessary   r �ss n   C h e c k   i f   q u a r a n t i n e   s t a t u s   i s   s e t ;   r e m o v e   i f   n e c e s s a r yp tut Q    Fvw�>v k    =xx yzy r    '{|{ I   %�=}�<
�= .sysoexecTEXT���     TEXT} b    !~~ b    ��� m    �� ��� : x a t t r   - p   c o m . a p p l e . q u a r a n t i n e� 1    �;
�; 
spac n     ��� 1     �:
�: 
strq� o    �9�9 0 filepath filePath�<  | o      �8�8 $0 quarantinestatus quarantineStatusz ��7� Z   ( =���6�5� >  ( +��� o   ( )�4�4 $0 quarantinestatus quarantineStatus� m   ) *�� ���  � I  . 9�3��2
�3 .sysoexecTEXT���     TEXT� b   . 5��� b   . 1��� m   . /�� ��� : x a t t r   - d   c o m . a p p l e . q u a r a n t i n e� 1   / 0�1
�1 
spac� n   1 4��� 1   2 4�0
�0 
strq� o   1 2�/�/ 0 filepath filePath�2  �6  �5  �7  w R      �.�-�,
�. .ascrerr ****      � ****�-  �,  �>  u ��� l  G G�+�*�)�+  �*  �)  � ��� l  G G�(���(  �   Change file permissions   � ��� 0   C h a n g e   f i l e   p e r m i s s i o n s� ��'� Q   G f���� k   J \�� ��� I  J Y�&��%
�& .sysoexecTEXT���     TEXT� b   J U��� b   J Q��� b   J O��� b   J M��� m   J K�� ��� 
 c h m o d� 1   K L�$
�$ 
spac� o   M N�#�#  0 newpermissions newPermissions� 1   O P�"
�" 
spac� n   Q T��� 1   R T�!
�! 
strq� o   Q R� �  0 filepath filePath�%  � ��� L   Z \�� m   Z [�
� boovtrue�  � R      ���
� .ascrerr ****      � ****�  �  � L   d f�� m   d e�
� boovfals�'  Y ��� l     ����  �  �  � ��� i     #��� I      ���� $0 comparemd5hashes CompareMD5Hashes� ��� o      �� 0 paramstring paramString�  �  � k     G�� ��� l     ����  � b \ This will check the file integrity of the downloaded template against the known good value.   � ��� �   T h i s   w i l l   c h e c k   t h e   f i l e   i n t e g r i t y   o f   t h e   d o w n l o a d e d   t e m p l a t e   a g a i n s t   t h e   k n o w n   g o o d   v a l u e .� ��� r     ��� I      ���� 0 splitstring SplitString� ��� o    �� 0 paramstring paramString� ��� m    �� ���  - , -�  �  � J      �� ��� o      �� 0 filepath filePath� ��� o      �� 0 	validhash 	validHash�  � ��� l   �
�	��
  �	  �  � ��� Z    '����� H    �� I    ���� 0 doesfileexist DoesFileExist� ��� o    �� 0 filepath filePath�  �  � L   ! #�� m   ! "�
� boovfals�  �  � ��� l  ( (� �����   ��  ��  � ���� Q   ( G���� k   + =�� ��� r   + 8��� l  + 6������ I  + 6�����
�� .sysoexecTEXT���     TEXT� b   + 2��� b   + .��� m   + ,�� ���  m d 5   - q� 1   , -��
�� 
spac� n   . 1��� 1   / 1��
�� 
strq� o   . /���� 0 filepath filePath��  ��  ��  � o      ���� 0 checkresult checkResult� ���� L   9 =�� =  9 <��� o   9 :���� 0 checkresult checkResult� o   : ;���� 0 	validhash 	validHash��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   E G�� m   E F��
�� boovfals��  � ��� l     ��������  ��  ��  � ��� i   $ '��� I      ������� 0 copyfile CopyFile� ���� o      ���� 0 	filepaths 	filePaths��  ��  � k     8�� ��� l     ��� ��  � _ Y Self-explanatory. Copy file from place A to place B. The original file will still exist.     � �   S e l f - e x p l a n a t o r y .   C o p y   f i l e   f r o m   p l a c e   A   t o   p l a c e   B .   T h e   o r i g i n a l   f i l e   w i l l   s t i l l   e x i s t .�  r      I      ������ 0 splitstring SplitString  o    ���� 0 	filepaths 	filePaths 	��	 m    

 �  - , -��  ��   J        o      ���� 0 
targetfile 
targetFile �� o      ���� "0 destinationfile destinationFile��   �� Q    8 k    .  I   +����
�� .sysoexecTEXT���     TEXT b    ' b    # b    ! b     m       �!!  c p 1    ��
�� 
spac l    "����" n     #$# 1     ��
�� 
strq$ o    ���� 0 
targetfile 
targetFile��  ��   1   ! "��
�� 
spac l  # &%����% n   # &&'& 1   $ &��
�� 
strq' o   # $���� "0 destinationfile destinationFile��  ��  ��   (��( L   , .)) m   , -��
�� boovtrue��   R      ������
�� .ascrerr ****      � ****��  ��   L   6 8** m   6 7��
�� boovfals��  � +,+ l     ��������  ��  ��  , -.- i   ( +/0/ I      ��1���� 00 createzipwithlocal7zip CreateZipWithLocal7Zip1 2��2 o      ���� 0 
zipcommand 
zipCommand��  ��  0 Q     3453 k    66 787 I   ��9��
�� .sysoexecTEXT���     TEXT9 o    ���� 0 
zipcommand 
zipCommand��  8 :��: L   	 ;; m   	 
<< �==  S u c c e s s��  4 R      ������
�� .ascrerr ****      � ****��  ��  5 L    >> o    ���� 0 errmsg errMsg. ?@? l     ��������  ��  ��  @ ABA i   , /CDC I      ��E���� <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiverE F��F o      ���� 0 paramstring paramString��  ��  D k     <GG HIH l     ��JK��  J q k Create a ZIP file of all the PDFs in the target folder. Makes it simpler for you to send them to your KTs.   K �LL �   C r e a t e   a   Z I P   f i l e   o f   a l l   t h e   P D F s   i n   t h e   t a r g e t   f o l d e r .   M a k e s   i t   s i m p l e r   f o r   y o u   t o   s e n d   t h e m   t o   y o u r   K T s .I MNM r     OPO I      ��Q���� 0 splitstring SplitStringQ RSR o    ���� 0 paramstring paramStringS T��T m    UU �VV  - , -��  ��  P J      WW XYX o      ���� 0 savepath savePathY Z��Z o      ���� 0 zippath zipPath��  N [��[ Q    <\]^\ k    2__ `a` I   /��b��
�� .sysoexecTEXT���     TEXTb b    +cdc b    )efe b    'ghg b    #iji b    !klk b    mnm m    oo �pp  c dn 1    ��
�� 
spacl n     qrq 1     ��
�� 
strqr o    ���� 0 savepath savePathj m   ! "ss �tt (   & &   / u s r / b i n / z i p   - j  h n   # &uvu 1   $ &��
�� 
strqv o   # $���� 0 zippath zipPathf 1   ' (��
�� 
spacd m   ) *ww �xx 
 * . p d f��  a y��y L   0 2zz m   0 1{{ �||  S u c c e s s��  ] R      ������
�� .ascrerr ****      � ****��  ��  ^ L   : <}} o   : ;���� 0 errmsg errMsg��  B ~~ l     ��������  ��  ��   ��� i   0 3��� I      ������� 0 
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
�� boovfals��  � ��� l     ��������  ��  ��  � ��� i   4 7��� I      ������� "0 doesbundleexist DoesBundleExist� ���� o      ���� 0 
bundlepath 
bundlePath��  ��  � k     �� ��� l     ������  � D > Used to check if the Dialog Toolkit Plus script bundle exists   � ��� |   U s e d   t o   c h e c k   i f   t h e   D i a l o g   T o o l k i t   P l u s   s c r i p t   b u n d l e   e x i s t s� ���� O    ��� L    �� l   ���~� I   �}��|
�} .coredoexnull���     ****� 4    �{�
�{ 
ditm� o    �z�z 0 
bundlepath 
bundlePath�|  �  �~  � m     ���                                                                                  sevs  alis    \  Macintosh HD               �=,�BD ����System Events.app                                              �����=,�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  ��  � ��� l     �y�x�w�y  �x  �w  � ��� i   8 ;��� I      �v��u�v 0 doesfileexist DoesFileExist� ��t� o      �s�s 0 filepath filePath�t  �u  � k     �� ��� l     �r���r  �   Self-explanatory   � ��� "   S e l f - e x p l a n a t o r y� ��q� O    ��� L    �� F    ��� l   ��p�o� I   �n��m
�n .coredoexnull���     ****� 4    �l�
�l 
ditm� o    �k�k 0 filepath filePath�m  �p  �o  � =    ��� n    ��� m    �j
�j 
pcls� 4    �i�
�i 
ditm� o    �h�h 0 filepath filePath� m    �g
�g 
file� m     ���                                                                                  sevs  alis    \  Macintosh HD               �=,�BD ����System Events.app                                              �����=,�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  �q  � ��� l     �f�e�d�f  �e  �d  � ��� i   < ?��� I      �c��b�c 0 downloadfile DownloadFile� ��a� o      �`�` 0 paramstring paramString�a  �b  � k     >�� ��� l     �_���_  � Z T Self-explanatory. The value of fileURL is the internet address to the desired file.   � ��� �   S e l f - e x p l a n a t o r y .   T h e   v a l u e   o f   f i l e U R L   i s   t h e   i n t e r n e t   a d d r e s s   t o   t h e   d e s i r e d   f i l e .� ��� r     ��� I      �^��]�^ 0 splitstring SplitString� ��� o    �\�\ 0 paramstring paramString� ��[� m    �� ���  - , -�[  �]  � J      �� ��� o      �Z�Z "0 destinationpath destinationPath� ��Y� o      �X�X 0 fileurl fileURL�Y  � ��W� Q    >���� k    ,�� ��� I   )�V��U
�V .sysoexecTEXT���     TEXT� b    %��� b    !��� b       m     �  c u r l   - L   - o   l   �T�S n     1    �R
�R 
strq o    �Q�Q "0 destinationpath destinationPath�T  �S  � m      �   � l  ! $	�P�O	 n   ! $

 1   " $�N
�N 
strq o   ! "�M�M 0 fileurl fileURL�P  �O  �U  � �L L   * , m   * +�K
�K boovtrue�L  � R      �J�I�H
�J .ascrerr ****      � ****�I  �H  � k   4 >  I  4 ;�G�F
�G .sysodlogaskr        TEXT b   4 7 m   4 5 � 0 E r r o r   d o w n l o a d i n g   f i l e :   o   5 6�E�E 0 fileurl fileURL�F   �D L   < > m   < =�C
�C boovfals�D  �W  �  l     �B�A�@�B  �A  �@    i   @ C I      �?�>�? 0 findsignature FindSignature �= o      �<�< 0 signaturepath signaturePath�=  �>   k     3   !"! l     �;#$�;  # m g If your signature isn't embedded in the Excel file, it will try to find an external JPG or PNG version   $ �%% �   I f   y o u r   s i g n a t u r e   i s n ' t   e m b e d d e d   i n   t h e   E x c e l   f i l e ,   i t   w i l l   t r y   t o   f i n d   a n   e x t e r n a l   J P G   o r   P N G   v e r s i o n" &�:& Q     3'()' Z    )*+,-* I    �9.�8�9 0 doesfileexist DoesFileExist. /�7/ b    010 o    �6�6 0 signaturepath signaturePath1 m    22 �33  m y S i g n a t u r e . p n g�7  �8  + L    44 b    565 o    �5�5 0 signaturepath signaturePath6 m    77 �88  m y S i g n a t u r e . p n g, 9:9 I    �4;�3�4 0 doesfileexist DoesFileExist; <�2< b    =>= o    �1�1 0 signaturepath signaturePath> m    ?? �@@  m y S i g n a t u r e . j p g�2  �3  : A�0A L     $BB b     #CDC o     !�/�/ 0 signaturepath signaturePathD m   ! "EE �FF  m y S i g n a t u r e . p n g�0  - L   ' )GG m   ' (HH �II  ( R      �.�-�,
�. .ascrerr ****      � ****�-  �,  ) L   1 3JJ m   1 2KK �LL  �:   MNM l     �+�*�)�+  �*  �)  N OPO i   D GQRQ I      �(S�'�( 0 installfonts InstallFontsS T�&T o      �%�% 0 paramstring paramString�&  �'  R k     QUU VWV r     XYX I      �$Z�#�$ 0 splitstring SplitStringZ [\[ o    �"�" 0 paramstring paramString\ ]�!] m    ^^ �__  - , -�!  �#  Y J      `` aba o      � �  0 fontname fontNameb c�c o      �� 0 fonturl fontURL�  W ded r    $fgf b    "hih b     jkj n    lml 1    �
� 
psxpm l   n��n I   �o�
� .earsffdralis        afdro m    �
� afdrcusr�  �  �  k m    pp �qq  L i b r a r y / F o n t s /i o     !�� 0 fontname fontNameg o      �� 0 userfontpath userFontPathe rsr r   % *tut b   % (vwv m   % &xx �yy  / L i b r a r y / F o n t s /w o   & '�� 0 fontname fontNameu o      ��  0 systemfontpath systemFontPaths z{z l  + +����  �  �  { |}| l  + +�~�  ~ U O Check if the font is already installed in user or system-wide font directories    ��� �   C h e c k   i f   t h e   f o n t   i s   a l r e a d y   i n s t a l l e d   i n   u s e r   o r   s y s t e m - w i d e   f o n t   d i r e c t o r i e s} ��� Z   + E����� G   + <��� I   + 1���� 0 doesfileexist DoesFileExist� ��� o   , -�
�
 0 userfontpath userFontPath�  �  � I   4 :�	���	 0 doesfileexist DoesFileExist� ��� o   5 6��  0 systemfontpath systemFontPath�  �  � L   ? A�� m   ? @�
� boovtrue�  �  � ��� l  F F����  �  �  � ��� l  F F����  � 2 , If not, download a copy to the fonts folder   � ��� X   I f   n o t ,   d o w n l o a d   a   c o p y   t o   t h e   f o n t s   f o l d e r� �� � L   F Q�� I   F P������� 0 downloadfile DownloadFile� ���� b   G L��� b   G J��� o   G H���� 0 userfontpath userFontPath� m   H I�� ���  - , -� o   J K���� 0 fonturl fontURL��  ��  �   P ��� l     ��������  ��  ��  � ��� i   H K��� I      ������� 0 
renamefile 
RenameFile� ���� o      ���� 0 paramstring paramString��  ��  � k     D�� ��� l     ������  � z t This pulls double duty for renaming a file or moving it to a new location. (It's the same process to the computer.)   � ��� �   T h i s   p u l l s   d o u b l e   d u t y   f o r   r e n a m i n g   a   f i l e   o r   m o v i n g   i t   t o   a   n e w   l o c a t i o n .   ( I t ' s   t h e   s a m e   p r o c e s s   t o   t h e   c o m p u t e r . )� ��� r     ��� I      ������� 0 splitstring SplitString� ��� o    ���� 0 paramstring paramString� ���� m    �� ���  - , -��  ��  � J      �� ��� o      ���� 0 
targetfile 
targetFile� ���� o      ���� 0 newfilename newFilename��  � ��� r    ��� n    ��� 1    ��
�� 
strq� n    ��� 1    ��
�� 
psxp� o    ���� 0 
targetfile 
targetFile� o      ���� 0 
targetfile 
targetFile� ��� r    &��� n    $��� 1   " $��
�� 
strq� n    "��� 1     "��
�� 
psxp� o     ���� 0 newfilename newFilename� o      ���� 0 newfilename newFilename� ���� Q   ' D���� k   * :�� ��� I  * 7�����
�� .sysoexecTEXT���     TEXT� b   * 3��� b   * 1��� b   * /��� b   * -��� m   * +�� ��� 
 m v   - f� 1   + ,��
�� 
spac� o   - .���� 0 
targetfile 
targetFile� 1   / 0��
�� 
spac� o   1 2���� 0 newfilename newFilename��  � ���� L   8 :�� m   8 9��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   B D�� m   B C��
�� boovfals��  � ��� l     ��������  ��  ��  � ��� i   L O��� I      ������� 0 savepptaspdf SavePptAsPdf� ���� o      ���� 0 tempsavepath tempSavePath��  ��  � Q     '���� k    �� ��� O    ��� k    �� ��� r    ��� 1    
��
�� 
AAPr� o      ���� 0 thisdocument thisDocument� ���� I   ����
�� .coresavenull���     obj � o    ���� 0 thisdocument thisDocument� ����
�� 
kfil� l   ������ 4    ���
�� 
psxf� o    ���� 0 tempsavepath tempSavePath��  ��  � �����
�� 
fltp� m    ��
�� pSAT � ��  ��  � m    ���                                                                                  PPT3  alis    Z  Macintosh HD               �=,�BD ����Microsoft PowerPoint.app                                       �����kd         ����  
 cu             Applications  (/:Applications:Microsoft PowerPoint.app/  2  M i c r o s o f t   P o w e r P o i n t . a p p    M a c i n t o s h   H D  %Applications/Microsoft PowerPoint.app   / ��  � ���� L    �� m    ��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   % '�� m   % &��
�� boovfals� � � l     ��������  ��  ��     l     ����     Folder Manipulation    � (   F o l d e r   M a n i p u l a t i o n  l     ��������  ��  ��   	 i   P S

 I      ������ 0 clearfolder ClearFolder �� o      ���� 0 foldertoempty folderToEmpty��  ��   k     ?  l     ����   h b Empties the target folder, but only of DOCX, PDF, and ZIP files. This folder will not be deleted.    � �   E m p t i e s   t h e   t a r g e t   f o l d e r ,   b u t   o n l y   o f   D O C X ,   P D F ,   a n d   Z I P   f i l e s .   T h i s   f o l d e r   w i l l   n o t   b e   d e l e t e d . �� Q     ? k    5  I   ����
�� .sysoexecTEXT���     TEXT b     b     b    
 !  b    "#" m    $$ �%%  f i n d# 1    ��
�� 
spac! l   	&����& n    	'(' 1    	��
�� 
strq( o    ���� 0 foldertoempty folderToEmpty��  ��   1   
 ��
�� 
spac m    )) �** : - t y p e   f   - n a m e   ' * . p d f '   - d e l e t e��   +,+ I   "��-��
�� .sysoexecTEXT���     TEXT- b    ./. b    010 b    232 b    454 m    66 �77  f i n d5 1    ��
�� 
spac3 l   8����8 n    9:9 1    ��
�� 
strq: o    ���� 0 foldertoempty folderToEmpty��  ��  1 1    ��
�� 
spac/ m    ;; �<< : - t y p e   f   - n a m e   ' * . z i p '   - d e l e t e��  , =>= I  # 2��?��
�� .sysoexecTEXT���     TEXT? b   # .@A@ b   # ,BCB b   # *DED b   # &FGF m   # $HH �II  f i n dG 1   $ %��
�� 
spacE l  & )J����J n   & )KLK 1   ' )��
�� 
strqL o   & '���� 0 foldertoempty folderToEmpty��  ��  C 1   * +��
�� 
spacA m   , -MM �NN < - t y p e   f   - n a m e   ' * . p p t x '   - d e l e t e��  > O��O L   3 5PP m   3 4��
�� boovtrue��   R      ������
�� .ascrerr ****      � ****��  ��   L   = ?QQ m   = >��
�� boovfals��  	 RSR l     ��������  ��  ��  S TUT i   T WVWV I      ��X���� .0 clearpdfsafterzipping ClearPDFsAfterZippingX Y��Y o      ���� 0 foldertoempty folderToEmpty��  ��  W Q     Z[\Z k    ]] ^_^ I   ��`��
�� .sysoexecTEXT���     TEXT` b    aba b    cdc b    
efe b    ghg m    ii �jj  f i n dh 1    ��
�� 
spacf l   	k����k n    	lml 1    	��
�� 
strqm o    ���� 0 foldertoempty folderToEmpty��  ��  d 1   
 ��
�� 
spacb m    nn �oo : - t y p e   f   - n a m e   ' * . p d f '   - d e l e t e��  _ p��p L    qq m    ��
�� boovtrue��  [ R      ����~
�� .ascrerr ****      � ****�  �~  \ L    rr m    �}
�} boovfalsU sts l     �|�{�z�|  �{  �z  t uvu i   X [wxw I      �yy�x�y 0 
copyfolder 
CopyFoldery z�wz o      �v�v 0 
folderpath 
folderPath�w  �x  x k     8{{ |}| l     �u~�u  ~ o i Self-explanatory. Copy a folder (or bundle) from place A to place B. The original file will still exist.    ��� �   S e l f - e x p l a n a t o r y .   C o p y   a   f o l d e r   ( o r   b u n d l e )   f r o m   p l a c e   A   t o   p l a c e   B .   T h e   o r i g i n a l   f i l e   w i l l   s t i l l   e x i s t .} ��� r     ��� I      �t��s�t 0 splitstring SplitString� ��� o    �r�r 0 
folderpath 
folderPath� ��q� m    �� ���  - , -�q  �s  � J      �� ��� o      �p�p 0 targetfolder targetFolder� ��o� o      �n�n &0 destinationfolder destinationFolder�o  � ��m� Q    8���� k    .�� ��� I   +�l��k
�l .sysoexecTEXT���     TEXT� b    '��� b    #��� b    !��� b    ��� m    �� ���  c p   - R f� 1    �j
�j 
spac� l    ��i�h� n     ��� 1     �g
�g 
strq� o    �f�f 0 targetfolder targetFolder�i  �h  � 1   ! "�e
�e 
spac� l  # &��d�c� n   # &��� 1   $ &�b
�b 
strq� o   # $�a�a &0 destinationfolder destinationFolder�d  �c  �k  � ��`� L   , .�� m   , -�_
�_ boovtrue�`  � R      �^�]�\
�^ .ascrerr ****      � ****�]  �\  � L   6 8�� m   6 7�[
�[ boovfals�m  v ��� l     �Z�Y�X�Z  �Y  �X  � ��� i   \ _��� I      �W��V�W 0 createfolder CreateFolder� ��U� o      �T�T 0 
folderpath 
folderPath�U  �V  � k     �� ��� l     �S���S  � \ V Self-explanatory. Needed for creating the folder for where the reports will be saved.   � ��� �   S e l f - e x p l a n a t o r y .   N e e d e d   f o r   c r e a t i n g   t h e   f o l d e r   f o r   w h e r e   t h e   r e p o r t s   w i l l   b e   s a v e d .� ��R� Q     ���� k    �� ��� I   �Q��P
�Q .sysoexecTEXT���     TEXT� b    
��� b    ��� m    �� ���  m k d i r   - p� 1    �O
�O 
spac� l   	��N�M� n    	��� 1    	�L
�L 
strq� o    �K�K 0 
folderpath 
folderPath�N  �M  �P  � ��J� L    �� m    �I
�I boovtrue�J  � R      �H�G�F
�H .ascrerr ****      � ****�G  �F  � L    �� m    �E
�E boovfals�R  � ��� l     �D�C�B�D  �C  �B  � ��� i   ` c��� I      �A��@�A 0 deletefolder DeleteFolder� ��?� o      �>�> 0 
folderpath 
folderPath�?  �@  � k     �� ��� l     �=���=  � c ] Self-explanatory. Same as with DeleteFile, extra security checks will likely be added later.   � ��� �   S e l f - e x p l a n a t o r y .   S a m e   a s   w i t h   D e l e t e F i l e ,   e x t r a   s e c u r i t y   c h e c k s   w i l l   l i k e l y   b e   a d d e d   l a t e r .� ��<� Q     ���� k    �� ��� I   �;��:
�; .sysoexecTEXT���     TEXT� b    
��� b    ��� m    �� ���  r m   - r f� 1    �9
�9 
spac� l   	��8�7� n    	��� 1    	�6
�6 
strq� o    �5�5 0 
folderpath 
folderPath�8  �7  �:  � ��4� L    �� m    �3
�3 boovtrue�4  � R      �2�1�0
�2 .ascrerr ****      � ****�1  �0  � L    �� m    �/
�/ boovfals�<  � ��� l     �.�-�,�.  �-  �,  � ��� i   d g��� I      �+��*�+ "0 doesfolderexist DoesFolderExist� ��)� o      �(�( 0 
folderpath 
folderPath�)  �*  � k     �� ��� l     �'���'  �   Self-explanatory   � ��� "   S e l f - e x p l a n a t o r y� ��&� O    ��� L    �� F       l   �%�$ I   �#�"
�# .coredoexnull���     **** 4    �!
�! 
ditm o    � �  0 
folderpath 
folderPath�"  �%  �$   =     n     m    �
� 
pcls 4    �	
� 
ditm	 o    �� 0 
folderpath 
folderPath m    �
� 
cfol� m     

�                                                                                  sevs  alis    \  Macintosh HD               �=,�BD ����System Events.app                                              �����=,�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  �&  �  l     ����  �  �    i   h k I      ��� (0 listfoldercontents ListFolderContents � o      �� 0 paramstring paramString�  �   k     i  r      I      ��� 0 splitstring SplitString  o    �� 0 paramstring paramString � m     �  - , -�  �   J         o      �� 0 
folderpath 
folderPath  !�! o      �� 0 fileextension fileExtension�   "#" l   ����  �  �  # $�
$ O    i%&% Q    h'()' k    Z** +,+ r    1-.- 6   //0/ n    &121 1   $ &�	
�	 
pnam2 n    $343 2  " $�
� 
file4 4    "�5
� 
cfol5 o     !�� 0 
folderpath 
folderPath0 =  ' .676 1   ( *�
� 
extn7 o   + -�� 0 fileextension fileExtension. o      �� 0 filelist fileList, 898 l  2 2��� �  �  �   9 :;: Z   2 ?<=����< =  2 6>?> o   2 3���� 0 filelist fileList? J   3 5����  = L   9 ;@@ m   9 :AA �BB  ��  ��  ; CDC l  @ @��������  ��  ��  D EFE r   @ EGHG n  @ CIJI 1   A C��
�� 
txdlJ 1   @ A��
�� 
ascrH o      ���� 00 oldtextitemsdelimiters oldTextItemsDelimitersF KLK r   F KMNM m   F GOO �PP  - , -N n     QRQ 1   H J��
�� 
txdlR 1   G H��
�� 
ascrL STS l  L L��������  ��  ��  T UVU r   L QWXW c   L OYZY o   L M���� 0 filelist fileListZ m   M N��
�� 
TEXTX o      ����  0 joinedfilelist joinedFileListV [\[ r   R W]^] o   R S���� 00 oldtextitemsdelimiters oldTextItemsDelimiters^ n     _`_ 1   T V��
�� 
txdl` 1   S T��
�� 
ascr\ aba l  X X��������  ��  ��  b c��c L   X Zdd o   X Y����  0 joinedfilelist joinedFileList��  ( R      ��e��
�� .ascrerr ****      � ****e o      ���� 0 errmsg errMsg��  ) L   b hff b   b gghg m   b eii �jj  E r r o r :  h o   e f���� 0 errmsg errMsg& m    kk�                                                                                  sevs  alis    \  Macintosh HD               �=,�BD ����System Events.app                                              �����=,�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  �
   lml l     ��������  ��  ��  m non i   l opqp I      ��r���� 0 
openfolder 
OpenFolderr s��s o      ���� 0 
folderpath 
folderPath��  ��  q Q     #tuvt k    ww xyx r    z{z c    	|}| 4    ��~
�� 
psxf~ o    ���� 0 
folderpath 
folderPath} m    ��
�� 
alis{ o      ���� 0 	pathalias 	pathAliasy �� O    ��� k    �� ��� I   �����
�� .aevtodocnull  �    alis� o    ���� 0 	pathalias 	pathAlias��  � ���� L    �� m    ��
�� boovtrue��  � m    ���                                                                                  MACS  alis    @  Macintosh HD               �=,�BD ����
Finder.app                                                     �����=,�        ����  
 cu             CoreServices  )/:System:Library:CoreServices:Finder.app/    
 F i n d e r . a p p    M a c i n t o s h   H D  &System/Library/CoreServices/Finder.app  / ��  ��  u R      ������
�� .ascrerr ****      � ****��  ��  v L   ! #�� m   ! "��
�� boovfalso ��� l     ��������  ��  ��  � ��� l     ������  �   Dialog Boxes   � ���    D i a l o g   B o x e s� ��� l     ��������  ��  ��  � ��� i   p s��� I      ������� 80 installdialogdisplayscript InstallDialogDisplayScript� ���� o      ���� 0 paramstring paramString��  ��  � k     �� ��� r     ��� b     	��� n     ��� 1    ��
�� 
psxp� l    ������ I    �����
�� .earsffdralis        afdr� m     ��
�� afdrcusr��  ��  ��  � m    �� ��� � L i b r a r y / A p p l i c a t i o n   S c r i p t s / c o m . m i c r o s o f t . E x c e l / D i a l o g D i s p l a y . s c p t� o      ���� 0 
scriptpath 
scriptPath� ��� r    ��� m    �� ��� � h t t p s : / / r a w . g i t h u b u s e r c o n t e n t . c o m / p a p e r c u t t e r 0 3 2 4 / S p e a k i n g E v a l s / m a i n / D i a l o g D i s p l a y . s c p t� o      ���� 0 downloadurl downloadURL� ��� l   ��������  ��  ��  � ��� l   ������  � A ; If an existing version is not found, download a fresh copy   � ��� v   I f   a n   e x i s t i n g   v e r s i o n   i s   n o t   f o u n d ,   d o w n l o a d   a   f r e s h   c o p y� ��� l   ������  � e _ Skip this first check until a full update function can be designed. For now, install each time   � ��� �   S k i p   t h i s   f i r s t   c h e c k   u n t i l   a   f u l l   u p d a t e   f u n c t i o n   c a n   b e   d e s i g n e d .   F o r   n o w ,   i n s t a l l   e a c h   t i m e� ��� l   ������  � 4 . if DoesFileExist(scriptPath) then return true   � ��� \   i f   D o e s F i l e E x i s t ( s c r i p t P a t h )   t h e n   r e t u r n   t r u e� ���� L    �� I    ������� 0 downloadfile DownloadFile� ���� b    ��� b    ��� o    ���� 0 
scriptpath 
scriptPath� m    �� ���  - , -� o    ���� 0 downloadurl downloadURL��  ��  ��  � ��� l     ��������  ��  ��  � ��� i   t w��� I      ������� >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder� ���� o      ���� 0 paramstring paramString��  ��  � k     ]�� ��� r     ��� b     	��� n     ��� 1    ��
�� 
psxp� l    ������ I    �����
�� .earsffdralis        afdr� m     ��
�� afdrcusr��  ��  ��  � m    �� ��� 0 L i b r a r y / S c r i p t   L i b r a r i e s� o      ���� .0 scriptlibrariesfolder scriptLibrariesFolder� ��� l   ��������  ��  ��  � ���� Z    ]������ I    ������� "0 doesfolderexist DoesFolderExist� ���� o    ���� .0 scriptlibrariesfolder scriptLibrariesFolder��  ��  � L    �� o    ���� .0 scriptlibrariesfolder scriptLibrariesFolder��  � Q    ]���� k    Q�� ��� l   ������  � m g ~/Library is typically a read-only folder, so I need to requst your password to create the need folder   � ��� �   ~ / L i b r a r y   i s   t y p i c a l l y   a   r e a d - o n l y   f o l d e r ,   s o   I   n e e d   t o   r e q u s t   y o u r   p a s s w o r d   t o   c r e a t e   t h e   n e e d   f o l d e r� ��� I   *����
�� .sysoexecTEXT���     TEXT� b    $��� b     ��� m    �� ���  m k d i r   - p� 1    ��
�� 
spac� n     #��� 1   ! #��
�� 
strq� o     !���� .0 scriptlibrariesfolder scriptLibrariesFolder� �����
�� 
badm� m   % &��
�� boovtrue��  � ��� l  + +������  � %  Set your username as the owner   � ��� >   S e t   y o u r   u s e r n a m e   a s   t h e   o w n e r�    I  + B��
�� .sysoexecTEXT���     TEXT b   + < b   + 8 b   + 6	 m   + ,

 �  c h o w n  	 n   , 5 1   3 5��
�� 
strq l  , 3���� n   , 3 1   1 3��
�� 
sisn l  , 1���� I  , 1������
�� .sysosigtsirr   ��� null��  ��  ��  ��  ��  ��   1   6 7��
�� 
spac n   8 ; 1   9 ;��
�� 
strq o   8 9���� .0 scriptlibrariesfolder scriptLibrariesFolder ��~
� 
badm m   = >�}
�} boovtrue�~    l  C C�|�|   5 / Give your username READ and WRITE permissions.    � ^   G i v e   y o u r   u s e r n a m e   R E A D   a n d   W R I T E   p e r m i s s i o n s .  I  C N�{
�{ .sysoexecTEXT���     TEXT b   C H m   C D   �!!  c h m o d   u + r w   n   D G"#" 1   E G�z
�z 
strq# o   D E�y�y .0 scriptlibrariesfolder scriptLibrariesFolder �x$�w
�x 
badm$ m   I J�v
�v boovtrue�w   %�u% L   O Q&& o   O P�t�t .0 scriptlibrariesfolder scriptLibrariesFolder�u  � R      �s�r�q
�s .ascrerr ****      � ****�r  �q  � L   Y ]'' m   Y \(( �))  ��  � *+* l     �p�o�n�p  �o  �n  + ,-, i   x {./. I      �m0�l�m 40 installdialogtoolkitplus InstallDialogToolkitPlus0 1�k1 o      �j�j "0 resourcesfolder resourcesFolder�k  �l  / k     �22 343 r     565 m     77 �88 � h t t p s : / / r a w . g i t h u b u s e r c o n t e n t . c o m / p a p e r c u t t e r 0 3 2 4 / S p e a k i n g E v a l s / m a i n / D i a l o g _ T o o l k i t . z i p6 o      �i�i 0 downloadurl downloadURL4 9:9 r    ;<; b    =>= n    ?@? 1   	 �h
�h 
psxp@ l   	A�g�fA I   	�eB�d
�e .earsffdralis        afdrB m    �c
�c afdrcusr�d  �g  �f  > m    CC �DD 0 L i b r a r y / S c r i p t   L i b r a r i e s< o      �b�b .0 scriptlibrariesfolder scriptLibrariesFolder: EFE r    GHG m    II �JJ 4 / D i a l o g   T o o l k i t   P l u s . s c p t dH o      �a�a $0 dialogbundlename dialogBundleNameF KLK r    MNM b    OPO o    �`�` .0 scriptlibrariesfolder scriptLibrariesFolderP o    �_�_ $0 dialogbundlename dialogBundleNameN o      �^�^ 20 dialogtoolkitplusbundle dialogToolkitPlusBundleL QRQ r    STS b    UVU o    �]�] "0 resourcesfolder resourcesFolderV m    WW �XX & / D i a l o g _ T o o l k i t . z i pT o      �\�\ 0 zipfilepath zipFilePathR YZY r     %[\[ b     #]^] o     !�[�[ "0 resourcesfolder resourcesFolder^ m   ! "__ �`` $ / d i a l o g T o o l k i t T e m p\ o      �Z�Z &0 zipextractionpath zipExtractionPathZ aba l  & &�Y�X�W�Y  �X  �W  b cdc l  & &�Vef�V  e 0 * Initial check to see if already installed   f �gg T   I n i t i a l   c h e c k   t o   s e e   i f   a l r e a d y   i n s t a l l e dd hih Z  & 5jk�U�Tj I   & ,�Sl�R�S "0 doesbundleexist DoesBundleExistl m�Qm o   ' (�P�P 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�Q  �R  k L   / 1nn m   / 0�O
�O boovtrue�U  �T  i opo l  6 6�N�M�L�N  �M  �L  p qrq l  6 6�Kst�K  s 3 - Ensure resources folder exists for later use   t �uu Z   E n s u r e   r e s o u r c e s   f o l d e r   e x i s t s   f o r   l a t e r   u s er vwv Z   6 Wxy�J�Ix H   6 =zz I   6 <�H{�G�H "0 doesfolderexist DoesFolderExist{ |�F| o   7 8�E�E "0 resourcesfolder resourcesFolder�F  �G  y Q   @ S}~} I   C I�D��C�D 0 createfolder CreateFolder� ��B� o   D E�A�A "0 resourcesfolder resourcesFolder�B  �C  ~ R      �@�?�>
�@ .ascrerr ****      � ****�?  �>   L   Q S�� m   Q R�=
�= boovfals�J  �I  w ��� l  X X�<�;�:�<  �;  �:  � ��� l  X X�9���9  � G A Check for a local copy and move it to the needed folder if found   � ��� �   C h e c k   f o r   a   l o c a l   c o p y   a n d   m o v e   i t   t o   t h e   n e e d e d   f o l d e r   i f   f o u n d� ��� Z   X |���8�7� I   X `�6��5�6 "0 doesbundleexist DoesBundleExist� ��4� b   Y \��� o   Y Z�3�3 "0 resourcesfolder resourcesFolder� o   Z [�2�2 $0 dialogbundlename dialogBundleName�4  �5  � Z   c x���1�0� I   c o�/��.�/ 0 
copyfolder 
CopyFolder� ��-� b   d k��� b   d i��� b   d g��� o   d e�,�, "0 resourcesfolder resourcesFolder� o   e f�+�+ $0 dialogbundlename dialogBundleName� m   g h�� ���  - , -� o   i j�*�* 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�-  �.  � L   r t�� m   r s�)
�) boovtrue�1  �0  �8  �7  � ��� l  } }�(�'�&�(  �'  �&  � ��� l  } }�%���%  � !  Otherwise, download and...   � ��� 6   O t h e r w i s e ,   d o w n l o a d   a n d . . .� ��� Z   } ����$�#� I   } ��"��!�" 0 downloadfile DownloadFile� �� � b   ~ ���� b   ~ ���� o   ~ �� 0 zipfilepath zipFilePath� m    ��� ���  - , -� o   � ��� 0 downloadurl downloadURL�   �!  � Q   � ����� k   � ��� ��� l  � �����  �   ...extract the files...   � ��� 0   . . . e x t r a c t   t h e   f i l e s . . .� ��� I  � ����
� .sysoexecTEXT���     TEXT� b   � ���� b   � ���� b   � ���� b   � ���� m   � ��� ���  u n z i p   - o� 1   � ��
� 
spac� l  � ����� n   � ���� 1   � ��
� 
strq� o   � ��� 0 zipfilepath zipFilePath�  �  � m   � ��� ���    - d  � l  � ����� n   � ���� 1   � ��
� 
strq� o   � ��� &0 zipextractionpath zipExtractionPath�  �  �  � ��� l  � �����  � 6 0 ...keep a local copy in the resources folder...   � ��� `   . . . k e e p   a   l o c a l   c o p y   i n   t h e   r e s o u r c e s   f o l d e r . . .� ��� I   � ����� 0 
copyfolder 
CopyFolder� ��� b   � ���� b   � ���� b   � ���� b   � ���� b   � ���� o   � ��� &0 zipextractionpath zipExtractionPath� m   � ��� ���  / D i a l o g _ T o o l k i t� o   � ��� $0 dialogbundlename dialogBundleName� m   � ��� ���  - , -� o   � ��
�
 "0 resourcesfolder resourcesFolder� o   � ��	�	 $0 dialogbundlename dialogBundleName�  �  � ��� l  � �����  � ; 5 ...and copy the script bundle to the required folder   � ��� j   . . . a n d   c o p y   t h e   s c r i p t   b u n d l e   t o   t h e   r e q u i r e d   f o l d e r� ��� I   � ����� 0 
copyfolder 
CopyFolder� ��� b   � ���� b   � ���� b   � ���� b   � ���� o   � ��� &0 zipextractionpath zipExtractionPath� m   � ��� ���  / D i a l o g _ T o o l k i t� o   � ��� $0 dialogbundlename dialogBundleName� m   � ��� ���  - , -� o   � ��� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�  �  �  � R      � ����
�  .ascrerr ****      � ****��  ��  �  �$  �#  � ��� l  � ���������  ��  ��  � ��� l  � �������  � D > Remove unneeded files and folders created during this process   � �   |   R e m o v e   u n n e e d e d   f i l e s   a n d   f o l d e r s   c r e a t e d   d u r i n g   t h i s   p r o c e s s�  I   � ������� 0 
deletefile 
DeleteFile �� o   � ����� 0 zipfilepath zipFilePath��  ��    I   � ������� 0 deletefolder DeleteFolder �� o   � ����� &0 zipextractionpath zipExtractionPath��  ��   	
	 l  � ���������  ��  ��  
  l  � �����   V P One final check to verify installation was successful and return true if it was    � �   O n e   f i n a l   c h e c k   t o   v e r i f y   i n s t a l l a t i o n   w a s   s u c c e s s f u l   a n d   r e t u r n   t r u e   i f   i t   w a s �� L   � � I   � ������� "0 doesbundleexist DoesBundleExist �� o   � ����� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  ��  -  l     ��������  ��  ��    i   |  I      ������ 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus �� o      ���� "0 resourcesfolder resourcesFolder��  ��   k     U  r       b     	!"! n     #$# 1    ��
�� 
psxp$ l    %����% I    ��&��
�� .earsffdralis        afdr& m     ��
�� afdrcusr��  ��  ��  " m    '' �(( d L i b r a r y / S c r i p t   L i b r a r i e s / D i a l o g   T o o l k i t   P l u s . s c p t d  o      ���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle )*) r    +,+ b    -.- o    ���� "0 resourcesfolder resourcesFolder. m    // �00 4 / D i a l o g   T o o l k i t   P l u s . s c p t d, o      ���� 0 	localcopy 	localCopy* 121 l   ��������  ��  ��  2 343 Z    R56��75 I    ��8���� "0 doesbundleexist DoesBundleExist8 9��9 o    ���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  6 Q    L:;<: k    A== >?> Z   6@A����@ H    %BB I    $��C���� "0 doesbundleexist DoesBundleExistC D��D o     ���� 0 	localcopy 	localCopy��  ��  A I   ( 2��E���� 0 
copyfolder 
CopyFolderE F��F b   ) .GHG b   ) ,IJI o   ) *���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundleJ m   * +KK �LL  - , -H o   , -���� 0 	localcopy 	localCopy��  ��  ��  ��  ? MNM I   7 =��O���� 0 deletefolder DeleteFolderO P��P o   8 9���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  N Q��Q r   > ARSR m   > ?��
�� boovtrueS o      ���� 0 removalresult removalResult��  ; R      ������
�� .ascrerr ****      � ****��  ��  < r   I LTUT m   I J��
�� boovfalsU o      ���� 0 removalresult removalResult��  7 r   O RVWV m   O P��
�� boovtrueW o      ���� 0 removalresult removalResult4 XYX l  S S��������  ��  ��  Y Z��Z L   S U[[ o   S T���� 0 removalresult removalResult��   \��\ l     ��������  ��  ��  ��       "��]^_`abcdefghijklmnopqrstuvwxyz{|}��  ]  ������������������������������������������������������������������ 00 getscriptversionnumber GetScriptVersionNumber�� "0 getmacosversion GetMacOSVersion�� 0 splitstring SplitString�� 0 
joinstring 
JoinString�� "0 loadapplication LoadApplication�� 0 isapploaded IsAppLoaded�� "0 closepowerpoint ClosePowerPoint�� .0 changefilepermissions ChangeFilePermissions�� $0 comparemd5hashes CompareMD5Hashes�� 0 copyfile CopyFile�� 00 createzipwithlocal7zip CreateZipWithLocal7Zip�� <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiver�� 0 
deletefile 
DeleteFile�� "0 doesbundleexist DoesBundleExist�� 0 doesfileexist DoesFileExist�� 0 downloadfile DownloadFile�� 0 findsignature FindSignature�� 0 installfonts InstallFonts�� 0 
renamefile 
RenameFile�� 0 savepptaspdf SavePptAsPdf�� 0 clearfolder ClearFolder�� .0 clearpdfsafterzipping ClearPDFsAfterZipping�� 0 
copyfolder 
CopyFolder�� 0 createfolder CreateFolder�� 0 deletefolder DeleteFolder�� "0 doesfolderexist DoesFolderExist�� (0 listfoldercontents ListFolderContents�� 0 
openfolder 
OpenFolder�� 80 installdialogdisplayscript InstallDialogDisplayScript�� >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder�� 40 installdialogtoolkitplus InstallDialogToolkitPlus�� 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus^ �� ����~���� 00 getscriptversionnumber GetScriptVersionNumber�� ����� �  ���� 0 paramstring paramString��  ~ ���� 0 paramstring paramString ���� 5 P�� �_ �� &���������� "0 getmacosversion GetMacOSVersion�� ����� �  ���� 0 paramstring paramString��  � ��~� 0 paramstring paramString�~ 0 	osversion 	osVersion�  8�}�|�{
�} .sysoexecTEXT���     TEXT�|  �{  ��  �j E�O�W X  h` �z H�y�x���w�z 0 splitstring SplitString�y �v��v �  �u�t�u &0 passedparamstring passedParamString�t (0 parameterseparator parameterSeparator�x  � �s�r�q�p�s &0 passedparamstring passedParamString�r (0 parameterseparator parameterSeparator�q 00 oldtextitemsdelimiters oldTextItemsDelimiters�p *0 separatedparameters separatedParameters� �o�n�m
�o 
ascr
�n 
txdl
�m 
citm�w  � ��,E�O���,FO��-E�O���,FUO�a �l v�k�j���i�l 0 
joinstring 
JoinString�k �h��h �  �g�f�g $0 passedparamarray passedParamArray�f (0 parameterseparator parameterSeparator�j  � �e�d�c�b�e $0 passedparamarray passedParamArray�d (0 parameterseparator parameterSeparator�c 00 oldtextitemsdelimiters oldTextItemsDelimiters�b $0 joinedparameters joinedParameters� �a�`�_
�a 
ascr
�` 
txdl
�_ 
TEXT�i  � ��,E�O���,FO��&E�O���,FUO�b �^ ��]�\���[�^ "0 loadapplication LoadApplication�] �Z��Z �  �Y�Y 0 appname appName�\  � �X�W�V�X 0 appname appName�W 0 errmsg errMsg�V 0 errnum errNum� 	�U�T ��S� ��R � �
�U 
capp
�T .miscactvnull��� ��� null�S 0 errmsg errMsg� �Q�P�O
�Q 
errn�P 0 errnum errNum�O  
�R 
spac�[ * *�/ *j UO�W X  ��%�%�%�%�%�%c �N ��M�L���K�N 0 isapploaded IsAppLoaded�M �J��J �  �I�I 0 appname appName�L  � �H�G�F�E�H 0 appname appName�G 0 
loadresult 
loadResult�F 0 errmsg errMsg�E 0 errnum errNum� �D�C�B � �A�
�D 
prcs
�C 
pnam
�B 
spac�A 0 errmsg errMsg� �@�?�>
�@ 
errn�? 0 errnum errNum�>  �K ; (� *�-�,� ��%�%E�Y 	��%�%E�UO�W X  �%�%�%�%�%d �=�<�;���:�= "0 closepowerpoint ClosePowerPoint�< �9��9 �  �8�8 0 paramstring paramString�;  � �7�6�7 0 paramstring paramString�6 0 closeresult closeResult� K�5�48?�3CG�2�1M
�5 
prcs
�4 
pnam
�3 .aevtquitnull��� ��� null�2  �1  �: 4 +� #*�-�,� � *j UO�E�Y �E�O�UW 	X  	�e �0[�/�.���-�0 .0 changefilepermissions ChangeFilePermissions�/ �,��, �  �+�+ 0 paramstring paramString�.  � �*�)�(�'�* 0 paramstring paramString�)  0 newpermissions newPermissions�( 0 filepath filePath�' $0 quarantinestatus quarantineStatus� g�&�%��$�#�"���!� ��& 0 splitstring SplitString
�% 
cobj
�$ 
spac
�# 
strq
�" .sysoexecTEXT���     TEXT�!  �   �- g*��l+ E[�k/E�Z[�l/E�ZO (��%��,%j E�O�� ��%��,%j Y hW X 	 
hO ��%�%�%��,%j OeW 	X 	 
ff �������� $0 comparemd5hashes CompareMD5Hashes� ��� �  �� 0 paramstring paramString�  � ����� 0 paramstring paramString� 0 filepath filePath� 0 	validhash 	validHash� 0 checkresult checkResult� 
����������� 0 splitstring SplitString
� 
cobj� 0 doesfileexist DoesFileExist
� 
spac
� 
strq
� .sysoexecTEXT���     TEXT�  �  � H*��l+ E[�k/E�Z[�l/E�ZO*�k+  fY hO ��%��,%j E�O�� W 	X  	fg �������
� 0 copyfile CopyFile� �	��	 �  �� 0 	filepaths 	filePaths�  � ���� 0 	filepaths 	filePaths� 0 
targetfile 
targetFile� "0 destinationfile destinationFile� 	
�� ��� ����� 0 splitstring SplitString
� 
cobj
� 
spac
� 
strq
�  .sysoexecTEXT���     TEXT��  ��  �
 9*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW 	X  fh ��0���������� 00 createzipwithlocal7zip CreateZipWithLocal7Zip�� ����� �  ���� 0 
zipcommand 
zipCommand��  � ������ 0 
zipcommand 
zipCommand�� 0 errmsg errMsg� ��<����
�� .sysoexecTEXT���     TEXT��  ��  ��  �j  O�W 	X  �i ��D���������� <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiver�� ����� �  ���� 0 paramstring paramString��  � ���������� 0 paramstring paramString�� 0 savepath savePath�� 0 zippath zipPath�� 0 errmsg errMsg� U����o����sw��{������ 0 splitstring SplitString
�� 
cobj
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� =*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%�%�%j O�W 	X 
 �j ������������� 0 
deletefile 
DeleteFile�� ����� �  ���� 0 filepath filePath��  � ���� 0 filepath filePath� �����������
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  ��  ��%��,%j OeW 	X  fk ������������� "0 doesbundleexist DoesBundleExist�� ����� �  ���� 0 
bundlepath 
bundlePath��  � ���� 0 
bundlepath 
bundlePath� �����
�� 
ditm
�� .coredoexnull���     ****�� � *�/j Ul ������������� 0 doesfileexist DoesFileExist�� ����� �  ���� 0 filepath filePath��  � ���� 0 filepath filePath� �����������
�� 
ditm
�� .coredoexnull���     ****
�� 
pcls
�� 
file
�� 
bool�� � *�/j 	 *�/�,� �&Um ������������� 0 downloadfile DownloadFile�� ����� �  ���� 0 paramstring paramString��  � �������� 0 paramstring paramString�� "0 destinationpath destinationPath�� 0 fileurl fileURL� ����������������� 0 splitstring SplitString
�� 
cobj
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  
�� .sysodlogaskr        TEXT�� ?*��l+ E[�k/E�Z[�l/E�ZO ��,%�%��,%j OeW X  �%j 
Ofn ������������ 0 findsignature FindSignature�� ����� �  ���� 0 signaturepath signaturePath��  � ���� 0 signaturepath signaturePath� 	2��7?EH����K�� 0 doesfileexist DoesFileExist��  ��  �� 4 +*��%k+  	��%Y *��%k+  	��%Y �W 	X  �o ��R���������� 0 installfonts InstallFonts�� ����� �  ���� 0 paramstring paramString��  � ������������ 0 paramstring paramString�� 0 fontname fontName�� 0 fonturl fontURL�� 0 userfontpath userFontPath��  0 systemfontpath systemFontPath� ^����������px��������� 0 splitstring SplitString
�� 
cobj
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp�� 0 doesfileexist DoesFileExist
�� 
bool�� 0 downloadfile DownloadFile�� R*��l+ E[�k/E�Z[�l/E�ZO�j �,�%�%E�O�%E�O*�k+ 
 
*�k+ �& eY hO*��%�%k+ p ������������� 0 
renamefile 
RenameFile�� ����� �  ���� 0 paramstring paramString��  � �������� 0 paramstring paramString�� 0 
targetfile 
targetFile�� 0 newfilename newFilename� 
�������������������� 0 splitstring SplitString
�� 
cobj
�� 
psxp
�� 
strq
�� 
spac
�� .sysoexecTEXT���     TEXT��  ��  �� E*��l+ E[�k/E�Z[�l/E�ZO��,�,E�O��,�,E�O ��%�%�%�%j OeW 	X  	fq ������������ 0 savepptaspdf SavePptAsPdf�� �~��~ �  �}�} 0 tempsavepath tempSavePath��  � �|�{�| 0 tempsavepath tempSavePath�{ 0 thisdocument thisDocument� 
��z�y�x�w�v�u�t�s�r
�z 
AAPr
�y 
kfil
�x 
psxf
�w 
fltp
�v pSAT � �u 
�t .coresavenull���     obj �s  �r  � ( � *�,E�O��*�/��� UOeW 	X  	fr �q�p�o���n�q 0 clearfolder ClearFolder�p �m��m �  �l�l 0 foldertoempty folderToEmpty�o  � �k�k 0 foldertoempty folderToEmpty� $�j�i)�h6;HM�g�f
�j 
spac
�i 
strq
�h .sysoexecTEXT���     TEXT�g  �f  �n @ 7��%��,%�%�%j O��%��,%�%�%j O��%��,%�%�%j OeW 	X 	 
fs �eW�d�c���b�e .0 clearpdfsafterzipping ClearPDFsAfterZipping�d �a��a �  �`�` 0 foldertoempty folderToEmpty�c  � �_�_ 0 foldertoempty folderToEmpty� i�^�]n�\�[�Z
�^ 
spac
�] 
strq
�\ .sysoexecTEXT���     TEXT�[  �Z  �b   ��%��,%�%�%j OeW 	X  ft �Yx�X�W���V�Y 0 
copyfolder 
CopyFolder�X �U��U �  �T�T 0 
folderpath 
folderPath�W  � �S�R�Q�S 0 
folderpath 
folderPath�R 0 targetfolder targetFolder�Q &0 destinationfolder destinationFolder� 	��P�O��N�M�L�K�J�P 0 splitstring SplitString
�O 
cobj
�N 
spac
�M 
strq
�L .sysoexecTEXT���     TEXT�K  �J  �V 9*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW 	X  fu �I��H�G���F�I 0 createfolder CreateFolder�H �E��E �  �D�D 0 
folderpath 
folderPath�G  � �C�C 0 
folderpath 
folderPath� ��B�A�@�?�>
�B 
spac
�A 
strq
�@ .sysoexecTEXT���     TEXT�?  �>  �F  ��%��,%j OeW 	X  fv �=��<�;���:�= 0 deletefolder DeleteFolder�< �9��9 �  �8�8 0 
folderpath 
folderPath�;  � �7�7 0 
folderpath 
folderPath� ��6�5�4�3�2
�6 
spac
�5 
strq
�4 .sysoexecTEXT���     TEXT�3  �2  �:  ��%��,%j OeW 	X  fw �1��0�/���.�1 "0 doesfolderexist DoesFolderExist�0 �-��- �  �,�, 0 
folderpath 
folderPath�/  � �+�+ 0 
folderpath 
folderPath� 
�*�)�(�'�&
�* 
ditm
�) .coredoexnull���     ****
�( 
pcls
�' 
cfol
�& 
bool�. � *�/j 	 *�/�,� �&Ux �%�$�#���"�% (0 listfoldercontents ListFolderContents�$ �!��! �  � �  0 paramstring paramString�#  � �������� 0 paramstring paramString� 0 
folderpath 
folderPath� 0 fileextension fileExtension� 0 filelist fileList� 00 oldtextitemsdelimiters oldTextItemsDelimiters�  0 joinedfilelist joinedFileList� 0 errmsg errMsg� ��k�����A��O���i� 0 splitstring SplitString
� 
cobj
� 
cfol
� 
file
� 
pnam�  
� 
extn
� 
ascr
� 
txdl
� 
TEXT� 0 errmsg errMsg�  �" j*��l+ E[�k/E�Z[�l/E�ZO� O A*�/�-�,�[�,\Z�81E�O�jv  �Y hO��,E�O���,FO��&E�O���,FO�W X  a �%Uy �q�����
� 0 
openfolder 
OpenFolder� �	��	 �  �� 0 
folderpath 
folderPath�  � ��� 0 
folderpath 
folderPath� 0 	pathalias 	pathAlias� ������
� 
psxf
� 
alis
� .aevtodocnull  �    alis�  �  �
 $ *�/�&E�O� 
�j OeUW 	X  fz � ����������  80 installdialogdisplayscript InstallDialogDisplayScript�� ����� �  ���� 0 paramstring paramString��  � �������� 0 paramstring paramString�� 0 
scriptpath 
scriptPath�� 0 downloadurl downloadURL� �����������
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp�� 0 downloadfile DownloadFile�� �j �,�%E�O�E�O*��%�%k+ { ������������� >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder�� ����� �  ���� 0 paramstring paramString��  � ������ 0 paramstring paramString�� .0 scriptlibrariesfolder scriptLibrariesFolder� ������������������
���� ����(
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
sisn��  ��  �� ^�j �,�%E�O*�k+  �Y E 9��%��,%�el 	O�*j �,�,%�%��,%�el 	O���,%�el 	O�W X  a | ��/���������� 40 installdialogtoolkitplus InstallDialogToolkitPlus�� ����� �  ���� "0 resourcesfolder resourcesFolder��  � ���������������� "0 resourcesfolder resourcesFolder�� 0 downloadurl downloadURL�� .0 scriptlibrariesfolder scriptLibrariesFolder�� $0 dialogbundlename dialogBundleName�� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�� 0 zipfilepath zipFilePath�� &0 zipextractionpath zipExtractionPath� 7������CIW_��������������������������������
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
W 	X  fY hO*��%k+  *��%�%�%k+  eY hY hO*��%�%k+  T Ha _ %�a ,%a %�a ,%j O*�a %�%a %�%�%k+ O*�a %�%a %�%k+ W X  hY hO*�k+ O*�k+ O*�k+ } ������������ 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus�� ����� �  ���� "0 resourcesfolder resourcesFolder��  � ���������� "0 resourcesfolder resourcesFolder�� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�� 0 	localcopy 	localCopy�� 0 removalresult removalResult� ������'/��K��������
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp�� "0 doesbundleexist DoesBundleExist�� 0 
copyfolder 
CopyFolder�� 0 deletefolder DeleteFolder��  ��  �� V�j �,�%E�O��%E�O*�k+  6 (*�k+  *��%�%k+ Y hO*�k+ OeE�W 
X 	 
fE�Y eE�O�ascr  ��ޭ