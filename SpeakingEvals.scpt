FasdUAS 1.101.10   ��   ��    k             l      ��  ��    � |
Helper Scripts for the DYB Speaking Evaluations Excel spreadsheet

Version: 1.4.0
Build:   20250310
Warren Feltmate
� 2025
     � 	 	 � 
 H e l p e r   S c r i p t s   f o r   t h e   D Y B   S p e a k i n g   E v a l u a t i o n s   E x c e l   s p r e a d s h e e t 
 
 V e r s i o n :   1 . 4 . 0 
 B u i l d :       2 0 2 5 0 3 1 0 
 W a r r e n   F e l t m a t e 
 �   2 0 2 5 
   
  
 l     ��������  ��  ��        l     ��  ��      Environment Variables     �   ,   E n v i r o n m e n t   V a r i a b l e s      l     ��������  ��  ��        i         I      �� ���� 00 getscriptversionnumber GetScriptVersionNumber   ��  o      ���� 0 paramstring paramString��  ��    k            l     ��  ��    ? 9- Use build number to determine if an update is available     �   r -   U s e   b u i l d   n u m b e r   t o   d e t e r m i n e   i f   a n   u p d a t e   i s   a v a i l a b l e   ��  L          m     ���� 4����     ! " ! l     ��������  ��  ��   "  # $ # i     % & % I      �� '���� "0 getmacosversion GetMacOSVersion '  (�� ( o      ���� 0 paramstring paramString��  ��   & k      ) )  * + * l     �� , -��   , ` Z Not currently used, but could be helpful if there are issues with older versions of MacOS    - � . . �   N o t   c u r r e n t l y   u s e d ,   b u t   c o u l d   b e   h e l p f u l   i f   t h e r e   a r e   i s s u e s   w i t h   o l d e r   v e r s i o n s   o f   M a c O S +  /�� / Q      0 1�� 0 k     2 2  3 4 3 r    
 5 6 5 I   �� 7��
�� .sysoexecTEXT���     TEXT 7 m     8 8 � 9 9 . s w _ v e r s   - p r o d u c t V e r s i o n��   6 o      ���� 0 	osversion 	osVersion 4  :�� : L     ; ; o    ���� 0 	osversion 	osVersion��   1 R      ������
�� .ascrerr ****      � ****��  ��  ��  ��   $  < = < l     ��������  ��  ��   =  > ? > l     �� @ A��   @   Parameter Manipulation    A � B B .   P a r a m e t e r   M a n i p u l a t i o n ?  C D C l     ��������  ��  ��   D  E F E i     G H G I      �� I���� 0 splitstring SplitString I  J K J o      ���� &0 passedparamstring passedParamString K  L�� L o      ���� (0 parameterseparator parameterSeparator��  ��   H k      M M  N O N l     �� P Q��   P d ^ Excel can only pass on parameter to this file. This makes it possible to split one into many.    Q � R R �   E x c e l   c a n   o n l y   p a s s   o n   p a r a m e t e r   t o   t h i s   f i l e .   T h i s   m a k e s   i t   p o s s i b l e   t o   s p l i t   o n e   i n t o   m a n y . O  S T S O      U V U k     W W  X Y X r    	 Z [ Z 1    ��
�� 
txdl [ o      ���� 00 oldtextitemsdelimiters oldTextItemsDelimiters Y  \ ] \ r   
  ^ _ ^ o   
 ���� (0 parameterseparator parameterSeparator _ 1    ��
�� 
txdl ]  ` a ` r     b c b n     d e d 2   ��
�� 
citm e o    ���� &0 passedparamstring passedParamString c o      ���� *0 separatedparameters separatedParameters a  f�� f r     g h g o    ���� 00 oldtextitemsdelimiters oldTextItemsDelimiters h 1    ��
�� 
txdl��   V 1     ��
�� 
ascr T  i�� i L     j j o    ���� *0 separatedparameters separatedParameters��   F  k l k l     ��������  ��  ��   l  m n m l     �� o p��   o    Application Manipulations    p � q q 4   A p p l i c a t i o n   M a n i p u l a t i o n s n  r s r l     ��������  ��  ��   s  t u t i     v w v I      �� x���� "0 loadapplication LoadApplication x  y�� y o      ���� 0 appname appName��  ��   w k     ) z z  { | { l     �� } ~��   } < 6 A simple function to tell the needed program to open.    ~ �   l   A   s i m p l e   f u n c t i o n   t o   t e l l   t h e   n e e d e d   p r o g r a m   t o   o p e n . |  ��� � Q     ) � � � � k     � �  � � � O    � � � I  
 ������
�� .miscactvnull��� ��� null��  ��   � 4    �� �
�� 
capp � o    ���� 0 appname appName �  ��� � L     � � m     � � � � �  ��   � R      �� � �
�� .ascrerr ****      � **** � o      ���� 0 errmsg errMsg � �� ���
�� 
errn � o      ���� 0 errnum errNum��   � L    ) � � b    ( � � � b    & � � � b    $ � � � b    " � � � b      � � � b     � � � m     � � � � �  E r r o r   l o a d i n g � 1    ��
�� 
spac � o    ���� 0 appname appName � m     ! � � � � �  :   � o   " #���� 0 errnum errNum � m   $ % � � � � �    -   � o   & '���� 0 errmsg errMsg��   u  � � � l     ��������  ��  ��   �  � � � i     � � � I      �� ����� 0 isapploaded IsAppLoaded �  ��� � o      ���� 0 appname appName��  ��   � k     : � �  � � � l     �� � ���   � N H This lets Excel check that the other program is open before continuing.    � � � � �   T h i s   l e t s   E x c e l   c h e c k   t h a t   t h e   o t h e r   p r o g r a m   i s   o p e n   b e f o r e   c o n t i n u i n g . �  ��� � Q     : � � � � k    & � �  � � � O    # � � � Z    " � ��� � � E     � � � l    ����� � n     � � � 1   
 ��
�� 
pnam � 2    
��
�� 
prcs��  ��   � o    ���� 0 appname appName � r     � � � b     � � � b     � � � o    ���� 0 appname appName � 1    ��
�� 
spac � m     � � � � �  i s   n o w   r u n n i n g . � o      ���� 0 
loadresult 
loadResult��   � r    " � � � b      � � � b     � � � m     � � � � �  E r r o r   o p e n i n g � 1    ��
�� 
spac � o    ���� 0 appname appName � o      ���� 0 
loadresult 
loadResult � m     � ��                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��   �  ��� � L   $ & � � o   $ %���� 0 
loadresult 
loadResult��   � R      �� � �
�� .ascrerr ****      � **** � o      ���� 0 errmsg errMsg � �� ���
�� 
errn � o      ���� 0 errnum errNum��   � L   . : � � b   . 9 � � � b   . 7 � � � b   . 5 � � � b   . 3 � � � b   . 1 � � � m   . / � � � � �  E r r o r   l o a d i n g   � o   / 0���� 0 appname appName � m   1 2 � � � � �  :   � o   3 4���� 0 errnum errNum � m   5 6 � � � � �    -   � o   7 8���� 0 errmsg errMsg��   �  � � � l     ��������  ��  ��   �  � � � i     � � � I      �� ����� "0 closepowerpoint ClosePowerPoint �  ��� � o      ���� 0 paramstring paramString��  ��   � k     3 � �  � � � l     �� � ���   � { u This will completely close MS PowerPoint, even from the Dock. This reduces the chances of errors on subsequent runs.    � � � � �   T h i s   w i l l   c o m p l e t e l y   c l o s e   M S   P o w e r P o i n t ,   e v e n   f r o m   t h e   D o c k .   T h i s   r e d u c e s   t h e   c h a n c e s   o f   e r r o r s   o n   s u b s e q u e n t   r u n s . �  ��� � Q     3 � � � � O    ) � � � k    ( � �  �  � Z    %� E     l   �~�} n     1   
 �|
�| 
pnam 2    
�{
�{ 
prcs�~  �}   m    		 �

 ( M i c r o s o f t   P o w e r P o i n t k      O    I   �z�y�x
�z .aevtquitnull��� ��� null�y  �x   m    �                                                                                  PPT3  alis    L  macOS                      ��'�BD ����Microsoft PowerPoint.app                                       �����Ώ�        ����  
 cu             Applications  (/:Applications:Microsoft PowerPoint.app/  2  M i c r o s o f t   P o w e r P o i n t . a p p    m a c O S  %Applications/Microsoft PowerPoint.app   / ��   �w r     m     � P P o w e r P o i n t   h a s   s u c c e s s f u l l y   b e e n   c l o s e d . o      �v�v 0 closeresult closeResult�w  �   r   " % m   " # � H P o w e r P o i n t   i s   n o t   c u r r e n t l y   r u n n i n g . o      �u�u 0 closeresult closeResult  �t L   & ( o   & '�s�s 0 closeresult closeResult�t   � m    �                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��   � R      �r�q�p
�r .ascrerr ****      � ****�q  �p   � L   1 3 m   1 2 � \ T h e r e   w a s   a n   e r r o r   t r y i n g   t o   c l o s e   P o w e r P o i n t .��   �  !  l     �o�n�m�o  �n  �m  ! "#" l     �l$%�l  $   File Manipulation   % �&& $   F i l e   M a n i p u l a t i o n# '(' l     �k�j�i�k  �j  �i  ( )*) i    +,+ I      �h-�g�h .0 changefilepermissions ChangeFilePermissions- .�f. o      �e�e 0 paramstring paramString�f  �g  , k     f// 010 r     232 I      �d4�c�d 0 splitstring SplitString4 565 o    �b�b 0 paramstring paramString6 7�a7 m    88 �99  - , -�a  �c  3 J      :: ;<; o      �`�`  0 newpermissions newPermissions< =�_= o      �^�^ 0 filepath filePath�_  1 >?> l   �]�\�[�]  �\  �[  ? @A@ l   �ZBC�Z  B = 7 Check if quarantine status is set; remove if necessary   C �DD n   C h e c k   i f   q u a r a n t i n e   s t a t u s   i s   s e t ;   r e m o v e   i f   n e c e s s a r yA EFE Q    FGH�YG k    =II JKJ r    'LML I   %�XN�W
�X .sysoexecTEXT���     TEXTN b    !OPO b    QRQ m    SS �TT : x a t t r   - p   c o m . a p p l e . q u a r a n t i n eR 1    �V
�V 
spacP n     UVU 1     �U
�U 
strqV o    �T�T 0 filepath filePath�W  M o      �S�S $0 quarantinestatus quarantineStatusK W�RW Z   ( =XY�Q�PX >  ( +Z[Z o   ( )�O�O $0 quarantinestatus quarantineStatus[ m   ) *\\ �]]  Y I  . 9�N^�M
�N .sysoexecTEXT���     TEXT^ b   . 5_`_ b   . 1aba m   . /cc �dd : x a t t r   - d   c o m . a p p l e . q u a r a n t i n eb 1   / 0�L
�L 
spac` n   1 4efe 1   2 4�K
�K 
strqf o   1 2�J�J 0 filepath filePath�M  �Q  �P  �R  H R      �I�H�G
�I .ascrerr ****      � ****�H  �G  �Y  F ghg l  G G�F�E�D�F  �E  �D  h iji l  G G�Ckl�C  k   Change file permissions   l �mm 0   C h a n g e   f i l e   p e r m i s s i o n sj n�Bn Q   G fopqo k   J \rr sts I  J Y�Au�@
�A .sysoexecTEXT���     TEXTu b   J Uvwv b   J Qxyx b   J Oz{z b   J M|}| m   J K~~ � 
 c h m o d} 1   K L�?
�? 
spac{ o   M N�>�>  0 newpermissions newPermissionsy 1   O P�=
�= 
spacw n   Q T��� 1   R T�<
�< 
strq� o   Q R�;�; 0 filepath filePath�@  t ��:� L   Z \�� m   Z [�9
�9 boovtrue�:  p R      �8�7�6
�8 .ascrerr ****      � ****�7  �6  q L   d f�� m   d e�5
�5 boovfals�B  * ��� l     �4�3�2�4  �3  �2  � ��� i    ��� I      �1��0�1 $0 comparemd5hashes CompareMD5Hashes� ��/� o      �.�. 0 paramstring paramString�/  �0  � k     G�� ��� l     �-���-  � b \ This will check the file integrity of the downloaded template against the known good value.   � ��� �   T h i s   w i l l   c h e c k   t h e   f i l e   i n t e g r i t y   o f   t h e   d o w n l o a d e d   t e m p l a t e   a g a i n s t   t h e   k n o w n   g o o d   v a l u e .� ��� r     ��� I      �,��+�, 0 splitstring SplitString� ��� o    �*�* 0 paramstring paramString� ��)� m    �� ���  - , -�)  �+  � J      �� ��� o      �(�( 0 filepath filePath� ��'� o      �&�& 0 	validhash 	validHash�'  � ��� l   �%�$�#�%  �$  �#  � ��� Z    '���"�!� H    �� I    � ���  0 doesfileexist DoesFileExist� ��� o    �� 0 filepath filePath�  �  � L   ! #�� m   ! "�
� boovfals�"  �!  � ��� l  ( (����  �  �  � ��� Q   ( G���� k   + =�� ��� r   + 8��� l  + 6���� I  + 6���
� .sysoexecTEXT���     TEXT� b   + 2��� b   + .��� m   + ,�� ���  m d 5   - q� 1   , -�
� 
spac� n   . 1��� 1   / 1�
� 
strq� o   . /�� 0 filepath filePath�  �  �  � o      �� 0 checkresult checkResult� ��� L   9 =�� =  9 <��� o   9 :�� 0 checkresult checkResult� o   : ;�� 0 	validhash 	validHash�  � R      ���

� .ascrerr ****      � ****�  �
  � L   E G�� m   E F�	
�	 boovfals�  � ��� l     ����  �  �  � ��� i     #��� I      ���� 0 copyfile CopyFile� ��� o      �� 0 	filepaths 	filePaths�  �  � k     8�� ��� l     ����  � _ Y Self-explanatory. Copy file from place A to place B. The original file will still exist.   � ��� �   S e l f - e x p l a n a t o r y .   C o p y   f i l e   f r o m   p l a c e   A   t o   p l a c e   B .   T h e   o r i g i n a l   f i l e   w i l l   s t i l l   e x i s t .� ��� r     ��� I      � ����  0 splitstring SplitString� ��� o    ���� 0 	filepaths 	filePaths� ���� m    �� ���  - , -��  ��  � J      �� ��� o      ���� 0 
targetfile 
targetFile� ���� o      ���� "0 destinationfile destinationFile��  � ���� Q    8���� k    .�� ��� I   +�����
�� .sysoexecTEXT���     TEXT� b    '��� b    #��� b    !��� b    ��� m    �� ���  c p� 1    ��
�� 
spac� l    ������ n     ��� 1     ��
�� 
strq� o    ���� 0 
targetfile 
targetFile��  ��  � 1   ! "��
�� 
spac� l  # &������ n   # &��� 1   $ &��
�� 
strq� o   # $���� "0 destinationfile destinationFile��  ��  ��  � ���� L   , .�� m   , -��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   6 8�� m   6 7��
�� boovfals��  � ��� l     ��������  ��  ��  � ��� i   $ '   I      ������ 00 createzipwithlocal7zip CreateZipWithLocal7Zip �� o      ���� 0 
zipcommand 
zipCommand��  ��   Q      k     	 I   ��
��
�� .sysoexecTEXT���     TEXT
 o    ���� 0 
zipcommand 
zipCommand��  	 �� L   	  m   	 
 �  S u c c e s s��   R      ������
�� .ascrerr ****      � ****��  ��   L     o    ���� 0 errmsg errMsg�  l     ��������  ��  ��    i   ( + I      ������ <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiver �� o      ���� 0 paramstring paramString��  ��   k     <  l     ����   q k Create a ZIP file of all the PDFs in the target folder. Makes it simpler for you to send them to your KTs.    � �   C r e a t e   a   Z I P   f i l e   o f   a l l   t h e   P D F s   i n   t h e   t a r g e t   f o l d e r .   M a k e s   i t   s i m p l e r   f o r   y o u   t o   s e n d   t h e m   t o   y o u r   K T s .  r      !  I      ��"���� 0 splitstring SplitString" #$# o    ���� 0 paramstring paramString$ %��% m    && �''  - , -��  ��  ! J      (( )*) o      ���� 0 savepath savePath* +��+ o      ���� 0 zippath zipPath��   ,��, Q    <-./- k    200 121 I   /��3��
�� .sysoexecTEXT���     TEXT3 b    +454 b    )676 b    '898 b    #:;: b    !<=< b    >?> m    @@ �AA  c d? 1    ��
�� 
spac= n     BCB 1     ��
�� 
strqC o    ���� 0 savepath savePath; m   ! "DD �EE (   & &   / u s r / b i n / z i p   - j  9 n   # &FGF 1   $ &��
�� 
strqG o   # $���� 0 zippath zipPath7 1   ' (��
�� 
spac5 m   ) *HH �II 
 * . p d f��  2 J��J L   0 2KK m   0 1LL �MM  S u c c e s s��  . R      ������
�� .ascrerr ****      � ****��  ��  / L   : <NN o   : ;���� 0 errmsg errMsg��   OPO l     ��������  ��  ��  P QRQ i   , /STS I      ��U���� 0 
deletefile 
DeleteFileU V��V o      ���� 0 filepath filePath��  ��  T k     WW XYX l     ��Z[��  Z M GSelf-explanatory. This will delete the target file, skipping the Trash.   [ �\\ � S e l f - e x p l a n a t o r y .   T h i s   w i l l   d e l e t e   t h e   t a r g e t   f i l e ,   s k i p p i n g   t h e   T r a s h .Y ]^] l      ��_`��  _ � � The value of filePath passed to this function is always carefully considered
	(and limited), but at a future point, I will likely add in some safety checks for extra security
	to prevent a dangerous value accidentally being sent to this function.
	   ` �aa�   T h e   v a l u e   o f   f i l e P a t h   p a s s e d   t o   t h i s   f u n c t i o n   i s   a l w a y s   c a r e f u l l y   c o n s i d e r e d 
 	 ( a n d   l i m i t e d ) ,   b u t   a t   a   f u t u r e   p o i n t ,   I   w i l l   l i k e l y   a d d   i n   s o m e   s a f e t y   c h e c k s   f o r   e x t r a   s e c u r i t y 
 	 t o   p r e v e n t   a   d a n g e r o u s   v a l u e   a c c i d e n t a l l y   b e i n g   s e n t   t o   t h i s   f u n c t i o n . 
 	^ b��b Q     cdec k    ff ghg I   ��i��
�� .sysoexecTEXT���     TEXTi b    
jkj b    lml m    nn �oo 
 r m   - fm 1    ��
�� 
spack l   	p����p n    	qrq 1    	��
�� 
strqr o    ���� 0 filepath filePath��  ��  ��  h s��s L    tt m    ��
�� boovtrue��  d R      ������
�� .ascrerr ****      � ****��  ��  e L    uu m    ��
�� boovfals��  R vwv l     ��������  ��  ��  w xyx i   0 3z{z I      ��|���� "0 doesbundleexist DoesBundleExist| }��} o      ���� 0 
bundlepath 
bundlePath��  ��  { k     ~~ � l     ������  � D > Used to check if the Dialog Toolkit Plus script bundle exists   � ��� |   U s e d   t o   c h e c k   i f   t h e   D i a l o g   T o o l k i t   P l u s   s c r i p t   b u n d l e   e x i s t s� ���� O    ��� L    �� l   ������ I   �����
�� .coredoexnull���     ****� 4    ���
�� 
ditm� o    ���� 0 
bundlepath 
bundlePath��  ��  ��  � m     ���                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  ��  y ��� l     ��������  ��  ��  � ��� i   4 7��� I      ������� 0 doesfileexist DoesFileExist� ���� o      ���� 0 filepath filePath��  ��  � k     �� ��� l     ������  �   Self-explanatory   � ��� "   S e l f - e x p l a n a t o r y� ���� O    ��� L    �� F    ��� l   ������ I   �����
�� .coredoexnull���     ****� 4    ���
�� 
ditm� o    ���� 0 filepath filePath��  ��  ��  � =    ��� n    ��� m    ��
�� 
pcls� 4    ���
�� 
ditm� o    ���� 0 filepath filePath� m    ��
�� 
file� m     ���                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  ��  � ��� l     �������  ��  �  � ��� i   8 ;��� I      �~��}�~ 0 downloadfile DownloadFile� ��|� o      �{�{ 0 paramstring paramString�|  �}  � k     B�� ��� l     �z���z  � Z T Self-explanatory. The value of fileURL is the internet address to the desired file.   � ��� �   S e l f - e x p l a n a t o r y .   T h e   v a l u e   o f   f i l e U R L   i s   t h e   i n t e r n e t   a d d r e s s   t o   t h e   d e s i r e d   f i l e .� ��� r     ��� I      �y��x�y 0 splitstring SplitString� ��� o    �w�w 0 paramstring paramString� ��v� m    �� ���  - , -�v  �x  � J      �� ��� o      �u�u "0 destinationpath destinationPath� ��t� o      �s�s 0 fileurl fileURL�t  � ��r� Q    B���� k    .�� ��� I   +�q��p
�q .sysoexecTEXT���     TEXT� b    '��� b    #��� b    !��� b    ��� m    �� ���  c u r l   - L   - o� 1    �o
�o 
spac� l    ��n�m� n     ��� 1     �l
�l 
strq� o    �k�k "0 destinationpath destinationPath�n  �m  � 1   ! "�j
�j 
spac� l  # &��i�h� n   # &��� 1   $ &�g
�g 
strq� o   # $�f�f 0 fileurl fileURL�i  �h  �p  � ��e� L   , .�� m   , -�d
�d boovtrue�e  � R      �c�b�a
�c .ascrerr ****      � ****�b  �a  � k   6 B�� ��� I  6 ?�`��_
�` .sysodlogaskr        TEXT� b   6 ;��� b   6 9��� m   6 7�� ��� . E r r o r   d o w n l o a d i n g   f i l e :� 1   7 8�^
�^ 
spac� o   9 :�]�] 0 fileurl fileURL�_  � ��\� L   @ B�� m   @ A�[
�[ boovfals�\  �r  � ��� l     �Z�Y�X�Z  �Y  �X  � ��� i   < ?��� I      �W��V�W 0 findsignature FindSignature� ��U� o      �T�T 0 signaturepath signaturePath�U  �V  � k     3�� ��� l     �S���S  � m g If your signature isn't embedded in the Excel file, it will try to find an external JPG or PNG version   � ��� �   I f   y o u r   s i g n a t u r e   i s n ' t   e m b e d d e d   i n   t h e   E x c e l   f i l e ,   i t   w i l l   t r y   t o   f i n d   a n   e x t e r n a l   J P G   o r   P N G   v e r s i o n� ��R� Q     3���� Z    )��� � I    �Q�P�Q 0 doesfileexist DoesFileExist �O b     o    �N�N 0 signaturepath signaturePath m     �  m y S i g n a t u r e . p n g�O  �P  � L     b    	 o    �M�M 0 signaturepath signaturePath	 m    

 �  m y S i g n a t u r e . p n g�  I    �L�K�L 0 doesfileexist DoesFileExist �J b     o    �I�I 0 signaturepath signaturePath m     �  m y S i g n a t u r e . j p g�J  �K   �H L     $ b     # o     !�G�G 0 signaturepath signaturePath m   ! " �  m y S i g n a t u r e . p n g�H    L   ' ) m   ' ( �  � R      �F�E�D
�F .ascrerr ****      � ****�E  �D  � L   1 3 m   1 2 �  �R  �  !  l     �C�B�A�C  �B  �A  ! "#" i   @ C$%$ I      �@&�?�@ 0 installfonts InstallFonts& '�>' o      �=�= 0 paramstring paramString�>  �?  % k     @(( )*) r     +,+ I      �<-�;�< 0 splitstring SplitString- ./. o    �:�: 0 paramstring paramString/ 0�90 m    11 �22  - , -�9  �;  , J      33 454 o      �8�8 0 fontname fontName5 6�76 o      �6�6 0 fonturl fontURL�7  * 787 r    $9:9 b    ";<; b     =>= n    ?@? 1    �5
�5 
psxp@ l   A�4�3A I   �2B�1
�2 .earsffdralis        afdrB m    �0
�0 afdrcusr�1  �4  �3  > m    CC �DD  L i b r a r y / F o n t s /< o     !�/�/ 0 fontname fontName: o      �.�. 0 fontpath fontPath8 EFE l  % %�-�,�+�-  �,  �+  F GHG l  % %�*IJ�*  I - ' Check if the font is already installed   J �KK N   C h e c k   i f   t h e   f o n t   i s   a l r e a d y   i n s t a l l e dH LML Z   % 4NO�)�(N I   % +�'P�&�' 0 doesfileexist DoesFileExistP Q�%Q o   & '�$�$ 0 fontpath fontPath�%  �&  O L   . 0RR m   . /�#
�# boovtrue�)  �(  M STS l  5 5�"�!� �"  �!  �   T UVU l  5 5�WX�  W 2 , If not, download a copy to the fonts folder   X �YY X   I f   n o t ,   d o w n l o a d   a   c o p y   t o   t h e   f o n t s   f o l d e rV Z�Z L   5 @[[ I   5 ?�\�� 0 downloadfile DownloadFile\ ]�] b   6 ;^_^ b   6 9`a` o   6 7�� 0 fontpath fontPatha m   7 8bb �cc  - , -_ o   9 :�� 0 fonturl fontURL�  �  �  # ded l     ����  �  �  e fgf i   D Ghih I      �j�� 0 
renamefile 
RenameFilej k�k o      �� 0 paramstring paramString�  �  i k     Dll mnm l     �op�  o z t This pulls double duty for renaming a file or moving it to a new location. (It's the same process to the computer.)   p �qq �   T h i s   p u l l s   d o u b l e   d u t y   f o r   r e n a m i n g   a   f i l e   o r   m o v i n g   i t   t o   a   n e w   l o c a t i o n .   ( I t ' s   t h e   s a m e   p r o c e s s   t o   t h e   c o m p u t e r . )n rsr r     tut I      �v�� 0 splitstring SplitStringv wxw o    �� 0 paramstring paramStringx y�y m    zz �{{  - , -�  �  u J      || }~} o      �� 0 
targetfile 
targetFile~ � o      �
�
 0 newfilename newFilename�  s ��� r    ��� n    ��� 1    �	
�	 
strq� n    ��� 1    �
� 
psxp� o    �� 0 
targetfile 
targetFile� o      �� 0 
targetfile 
targetFile� ��� r    &��� n    $��� 1   " $�
� 
strq� n    "��� 1     "�
� 
psxp� o     �� 0 newfilename newFilename� o      �� 0 newfilename newFilename� ��� Q   ' D���� k   * :�� ��� I  * 7� ���
�  .sysoexecTEXT���     TEXT� b   * 3��� b   * 1��� b   * /��� b   * -��� m   * +�� ��� 
 m v   - f� 1   + ,��
�� 
spac� o   - .���� 0 
targetfile 
targetFile� 1   / 0��
�� 
spac� o   1 2���� 0 newfilename newFilename��  � ���� L   8 :�� m   8 9��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   B D�� m   B C��
�� boovfals�  g ��� l     ��������  ��  ��  � ��� i   H K��� I      ������� 0 savepptaspdf SavePptAsPdf� ���� o      ���� 0 tempsavepath tempSavePath��  ��  � Q     '���� k    �� ��� O    ��� k    �� ��� r    ��� 1    
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
�� pSAT � ��  ��  � m    ���                                                                                  PPT3  alis    L  macOS                      ��'�BD ����Microsoft PowerPoint.app                                       �����Ώ�        ����  
 cu             Applications  (/:Applications:Microsoft PowerPoint.app/  2  M i c r o s o f t   P o w e r P o i n t . a p p    m a c O S  %Applications/Microsoft PowerPoint.app   / ��  � ���� L    �� m    ��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   % '�� m   % &��
�� boovfals� ��� l     ��������  ��  ��  � ��� l     ������  �   Folder Manipulation   � ��� (   F o l d e r   M a n i p u l a t i o n� ��� l     ��������  ��  ��  � ��� i   L O��� I      ������� 0 clearfolder ClearFolder� ���� o      ���� 0 foldertoempty folderToEmpty��  ��  � k     ?�� ��� l     ������  � h b Empties the target folder, but only of DOCX, PDF, and ZIP files. This folder will not be deleted.   � ��� �   E m p t i e s   t h e   t a r g e t   f o l d e r ,   b u t   o n l y   o f   D O C X ,   P D F ,   a n d   Z I P   f i l e s .   T h i s   f o l d e r   w i l l   n o t   b e   d e l e t e d .� ���� Q     ?���� k    5�� ��� I   �����
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
spac� l   ������ n       1    ��
�� 
strq o    ���� 0 foldertoempty folderToEmpty��  ��  � 1    ��
�� 
spac� m     � : - t y p e   f   - n a m e   ' * . z i p '   - d e l e t e��  �  I  # 2����
�� .sysoexecTEXT���     TEXT b   # . b   # ,	
	 b   # * b   # & m   # $ �  f i n d 1   $ %��
�� 
spac l  & )���� n   & ) 1   ' )��
�� 
strq o   & '���� 0 foldertoempty folderToEmpty��  ��  
 1   * +��
�� 
spac m   , - � < - t y p e   f   - n a m e   ' * . p p t x '   - d e l e t e��   �� L   3 5 m   3 4��
�� boovtrue��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   = ? m   = >��
�� boovfals��  �  l     ��������  ��  ��    i   P S I      ������ .0 clearpdfsafterzipping ClearPDFsAfterZipping  ��  o      ���� 0 foldertoempty folderToEmpty��  ��   Q     !"#! k    $$ %&% I   ��'��
�� .sysoexecTEXT���     TEXT' b    ()( b    *+* b    
,-, b    ./. m    00 �11  f i n d/ 1    ��
�� 
spac- l   	2����2 n    	343 1    	��
�� 
strq4 o    ���� 0 foldertoempty folderToEmpty��  ��  + 1   
 ��
�� 
spac) m    55 �66 : - t y p e   f   - n a m e   ' * . p d f '   - d e l e t e��  & 7��7 L    88 m    ��
�� boovtrue��  " R      ������
�� .ascrerr ****      � ****��  ��  # L    99 m    ��
�� boovfals :;: l     ��������  ��  ��  ; <=< i   T W>?> I      ��@���� 0 
copyfolder 
CopyFolder@ A��A o      ���� 0 
folderpath 
folderPath��  ��  ? k     8BB CDC l     ��EF��  E o i Self-explanatory. Copy a folder (or bundle) from place A to place B. The original file will still exist.   F �GG �   S e l f - e x p l a n a t o r y .   C o p y   a   f o l d e r   ( o r   b u n d l e )   f r o m   p l a c e   A   t o   p l a c e   B .   T h e   o r i g i n a l   f i l e   w i l l   s t i l l   e x i s t .D HIH r     JKJ I      ��L���� 0 splitstring SplitStringL MNM o    ���� 0 
folderpath 
folderPathN O��O m    PP �QQ  - , -��  ��  K J      RR STS o      ���� 0 targetfolder targetFolderT U��U o      ���� &0 destinationfolder destinationFolder��  I V��V Q    8WXYW k    .ZZ [\[ I   +��]��
�� .sysoexecTEXT���     TEXT] b    '^_^ b    #`a` b    !bcb b    ded m    ff �gg  c p   - R fe 1    ��
�� 
spacc l    h����h n     iji 1     ��
�� 
strqj o    ���� 0 targetfolder targetFolder��  ��  a 1   ! "��
�� 
spac_ l  # &k����k n   # &lml 1   $ &��
�� 
strqm o   # $�� &0 destinationfolder destinationFolder��  ��  ��  \ n�~n L   , .oo m   , -�}
�} boovtrue�~  X R      �|�{�z
�| .ascrerr ****      � ****�{  �z  Y L   6 8pp m   6 7�y
�y boovfals��  = qrq l     �x�w�v�x  �w  �v  r sts i   X [uvu I      �uw�t�u 0 createfolder CreateFolderw x�sx o      �r�r 0 
folderpath 
folderPath�s  �t  v k     yy z{z l     �q|}�q  | \ V Self-explanatory. Needed for creating the folder for where the reports will be saved.   } �~~ �   S e l f - e x p l a n a t o r y .   N e e d e d   f o r   c r e a t i n g   t h e   f o l d e r   f o r   w h e r e   t h e   r e p o r t s   w i l l   b e   s a v e d .{ �p Q     ���� k    �� ��� I   �o��n
�o .sysoexecTEXT���     TEXT� b    
��� b    ��� m    �� ���  m k d i r   - p� 1    �m
�m 
spac� l   	��l�k� n    	��� 1    	�j
�j 
strq� o    �i�i 0 
folderpath 
folderPath�l  �k  �n  � ��h� L    �� m    �g
�g boovtrue�h  � R      �f�e�d
�f .ascrerr ****      � ****�e  �d  � L    �� m    �c
�c boovfals�p  t ��� l     �b�a�`�b  �a  �`  � ��� i   \ _��� I      �_��^�_ 0 deletefolder DeleteFolder� ��]� o      �\�\ 0 
folderpath 
folderPath�]  �^  � k     �� ��� l     �[���[  � c ] Self-explanatory. Same as with DeleteFile, extra security checks will likely be added later.   � ��� �   S e l f - e x p l a n a t o r y .   S a m e   a s   w i t h   D e l e t e F i l e ,   e x t r a   s e c u r i t y   c h e c k s   w i l l   l i k e l y   b e   a d d e d   l a t e r .� ��Z� Q     ���� k    �� ��� I   �Y��X
�Y .sysoexecTEXT���     TEXT� b    
��� b    ��� m    �� ���  r m   - r f� 1    �W
�W 
spac� l   	��V�U� n    	��� 1    	�T
�T 
strq� o    �S�S 0 
folderpath 
folderPath�V  �U  �X  � ��R� L    �� m    �Q
�Q boovtrue�R  � R      �P�O�N
�P .ascrerr ****      � ****�O  �N  � L    �� m    �M
�M boovfals�Z  � ��� l     �L�K�J�L  �K  �J  � ��� i   ` c��� I      �I��H�I "0 doesfolderexist DoesFolderExist� ��G� o      �F�F 0 
folderpath 
folderPath�G  �H  � k     �� ��� l     �E���E  �   Self-explanatory   � ��� "   S e l f - e x p l a n a t o r y� ��D� O    ��� L    �� F    ��� l   ��C�B� I   �A��@
�A .coredoexnull���     ****� 4    �?�
�? 
ditm� o    �>�> 0 
folderpath 
folderPath�@  �C  �B  � =    ��� n    ��� m    �=
�= 
pcls� 4    �<�
�< 
ditm� o    �;�; 0 
folderpath 
folderPath� m    �:
�: 
cfol� m     ���                                                                                  sevs  alis    N  macOS                      ��'�BD ����System Events.app                                              ������'�        ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    m a c O S  -System/Library/CoreServices/System Events.app   / ��  �D  � ��� l     �9�8�7�9  �8  �7  � ��� l     �6���6  �   Dialog Boxes   � ���    D i a l o g   B o x e s� ��� l     �5�4�3�5  �4  �3  � ��� i   d g��� I      �2��1�2 80 installdialogdisplayscript InstallDialogDisplayScript� ��0� o      �/�/ 0 paramstring paramString�0  �1  � k     �� ��� r     ��� b     	��� n     ��� 1    �.
�. 
psxp� l    ��-�,� I    �+��*
�+ .earsffdralis        afdr� m     �)
�) afdrcusr�*  �-  �,  � m    �� ��� � L i b r a r y / A p p l i c a t i o n   S c r i p t s / c o m . m i c r o s o f t . E x c e l / D i a l o g D i s p l a y . s c p t� o      �(�( 0 
scriptpath 
scriptPath� ��� r    ��� m    �� ��� � h t t p s : / / r a w . g i t h u b u s e r c o n t e n t . c o m / p a p e r c u t t e r 0 3 2 4 / S p e a k i n g E v a l s / m a i n / D i a l o g D i s p l a y . s c p t� o      �'�' 0 downloadurl downloadURL� ��� l   �&�%�$�&  �%  �$  � ��� l   �#���#  � A ; If an existing version is not found, download a fresh copy   � ��� v   I f   a n   e x i s t i n g   v e r s i o n   i s   n o t   f o u n d ,   d o w n l o a d   a   f r e s h   c o p y� ��� l   �"���"  � e _ Skip this first check until a full update function can be designed. For now, install each time   � ��� �   S k i p   t h i s   f i r s t   c h e c k   u n t i l   a   f u l l   u p d a t e   f u n c t i o n   c a n   b e   d e s i g n e d .   F o r   n o w ,   i n s t a l l   e a c h   t i m e�    l   �!�!   4 . if DoesFileExist(scriptPath) then return true    � \   i f   D o e s F i l e E x i s t ( s c r i p t P a t h )   t h e n   r e t u r n   t r u e �  L     I    ��� 0 downloadfile DownloadFile � b    	
	 b     o    �� 0 
scriptpath 
scriptPath m     �  - , -
 o    �� 0 downloadurl downloadURL�  �  �   �  l     ����  �  �    i   h k I      ��� >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder � o      �� 0 paramstring paramString�  �   k     ]  r      b     	 n      1    �
� 
psxp l     ��  I    �!�
� .earsffdralis        afdr! m     �
� afdrcusr�  �  �   m    "" �## 0 L i b r a r y / S c r i p t   L i b r a r i e s o      �� .0 scriptlibrariesfolder scriptLibrariesFolder $%$ l   ���
�  �  �
  % &�	& Z    ]'(�)' I    �*�� "0 doesfolderexist DoesFolderExist* +�+ o    �� .0 scriptlibrariesfolder scriptLibrariesFolder�  �  ( L    ,, o    �� .0 scriptlibrariesfolder scriptLibrariesFolder�  ) Q    ]-./- k    Q00 121 l   �34�  3 m g ~/Library is typically a read-only folder, so I need to requst your password to create the need folder   4 �55 �   ~ / L i b r a r y   i s   t y p i c a l l y   a   r e a d - o n l y   f o l d e r ,   s o   I   n e e d   t o   r e q u s t   y o u r   p a s s w o r d   t o   c r e a t e   t h e   n e e d   f o l d e r2 676 I   *�89
� .sysoexecTEXT���     TEXT8 b    $:;: b     <=< m    >> �??  m k d i r   - p= 1    � 
�  
spac; n     #@A@ 1   ! #��
�� 
strqA o     !���� .0 scriptlibrariesfolder scriptLibrariesFolder9 ��B��
�� 
badmB m   % &��
�� boovtrue��  7 CDC l  + +��EF��  E %  Set your username as the owner   F �GG >   S e t   y o u r   u s e r n a m e   a s   t h e   o w n e rD HIH I  + B��JK
�� .sysoexecTEXT���     TEXTJ b   + <LML b   + 8NON b   + 6PQP m   + ,RR �SS  c h o w n  Q n   , 5TUT 1   3 5��
�� 
strqU l  , 3V����V n   , 3WXW 1   1 3��
�� 
sisnX l  , 1Y����Y I  , 1������
�� .sysosigtsirr   ��� null��  ��  ��  ��  ��  ��  O 1   6 7��
�� 
spacM n   8 ;Z[Z 1   9 ;��
�� 
strq[ o   8 9���� .0 scriptlibrariesfolder scriptLibrariesFolderK ��\��
�� 
badm\ m   = >��
�� boovtrue��  I ]^] l  C C��_`��  _ 5 / Give your username READ and WRITE permissions.   ` �aa ^   G i v e   y o u r   u s e r n a m e   R E A D   a n d   W R I T E   p e r m i s s i o n s .^ bcb I  C N��de
�� .sysoexecTEXT���     TEXTd b   C Hfgf m   C Dhh �ii  c h m o d   u + r w  g n   D Gjkj 1   E G��
�� 
strqk o   D E���� .0 scriptlibrariesfolder scriptLibrariesFoldere ��l��
�� 
badml m   I J��
�� boovtrue��  c m��m L   O Qnn o   O P���� .0 scriptlibrariesfolder scriptLibrariesFolder��  . R      ������
�� .ascrerr ****      � ****��  ��  / L   Y ]oo m   Y \pp �qq  �	   rsr l     ��������  ��  ��  s tut i   l ovwv I      ��x���� 40 installdialogtoolkitplus InstallDialogToolkitPlusx y��y o      ���� "0 resourcesfolder resourcesFolder��  ��  w k     �zz {|{ r     }~} m      ��� � h t t p s : / / r a w . g i t h u b u s e r c o n t e n t . c o m / p a p e r c u t t e r 0 3 2 4 / S p e a k i n g E v a l s / m a i n / D i a l o g _ T o o l k i t . z i p~ o      ���� 0 downloadurl downloadURL| ��� r    ��� b    ��� n    ��� 1   	 ��
�� 
psxp� l   	������ I   	�����
�� .earsffdralis        afdr� m    ��
�� afdrcusr��  ��  ��  � m    �� ��� 0 L i b r a r y / S c r i p t   L i b r a r i e s� o      ���� .0 scriptlibrariesfolder scriptLibrariesFolder� ��� r    ��� m    �� ��� 4 / D i a l o g   T o o l k i t   P l u s . s c p t d� o      ���� $0 dialogbundlename dialogBundleName� ��� r    ��� b    ��� o    ���� .0 scriptlibrariesfolder scriptLibrariesFolder� o    ���� $0 dialogbundlename dialogBundleName� o      ���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle� ��� r    ��� b    ��� o    ���� "0 resourcesfolder resourcesFolder� m    �� ��� & / D i a l o g _ T o o l k i t . z i p� o      ���� 0 zipfilepath zipFilePath� ��� r     %��� b     #��� o     !���� "0 resourcesfolder resourcesFolder� m   ! "�� ��� $ / d i a l o g T o o l k i t T e m p� o      ���� &0 zipextractionpath zipExtractionPath� ��� l  & &��������  ��  ��  � ��� l  & &������  � 0 * Initial check to see if already installed   � ��� T   I n i t i a l   c h e c k   t o   s e e   i f   a l r e a d y   i n s t a l l e d� ��� Z  & 5������� I   & ,������� "0 doesbundleexist DoesBundleExist� ���� o   ' (���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  � L   / 1�� m   / 0��
�� boovtrue��  ��  � ��� l  6 6��������  ��  ��  � ��� l  6 6������  � 3 - Ensure resources folder exists for later use   � ��� Z   E n s u r e   r e s o u r c e s   f o l d e r   e x i s t s   f o r   l a t e r   u s e� ��� Z   6 W������� H   6 =�� I   6 <������� "0 doesfolderexist DoesFolderExist� ���� o   7 8���� "0 resourcesfolder resourcesFolder��  ��  � Q   @ S���� I   C I������� 0 createfolder CreateFolder� ���� o   D E���� "0 resourcesfolder resourcesFolder��  ��  � R      ������
�� .ascrerr ****      � ****��  ��  � L   Q S�� m   Q R��
�� boovfals��  ��  � ��� l  X X��������  ��  ��  � ��� l  X X������  � G A Check for a local copy and move it to the needed folder if found   � ��� �   C h e c k   f o r   a   l o c a l   c o p y   a n d   m o v e   i t   t o   t h e   n e e d e d   f o l d e r   i f   f o u n d� ��� Z   X |������� I   X `������� "0 doesbundleexist DoesBundleExist� ���� b   Y \��� o   Y Z���� "0 resourcesfolder resourcesFolder� o   Z [���� $0 dialogbundlename dialogBundleName��  ��  � Z   c x������� I   c o������� 0 
copyfolder 
CopyFolder� ���� b   d k��� b   d i��� b   d g��� o   d e���� "0 resourcesfolder resourcesFolder� o   e f���� $0 dialogbundlename dialogBundleName� m   g h�� ���  - , -� o   i j���� 20 dialogtoolkitplusbundle dialogToolkitPlusBundle��  ��  � L   r t�� m   r s��
�� boovtrue��  ��  ��  ��  � ��� l  } }��������  ��  ��  � ��� l  } }������  � !  Otherwise, download and...   � ��� 6   O t h e r w i s e ,   d o w n l o a d   a n d . . .� ��� Z   } �������� I   } �������� 0 downloadfile DownloadFile� ���� b   ~ ���� b   ~ ���� o   ~ ���� 0 zipfilepath zipFilePath� m    ��� ���  - , -� o   � ����� 0 downloadurl downloadURL��  ��  � Q   � ������ k   � ��� ��� l  � ���� ��  �   ...extract the files...     � 0   . . . e x t r a c t   t h e   f i l e s . . .�  I  � �����
�� .sysoexecTEXT���     TEXT b   � � b   � � b   � �	
	 b   � � m   � � �  u n z i p   - o 1   � ���
�� 
spac
 l  � ����� n   � � 1   � ���
�� 
strq o   � ����� 0 zipfilepath zipFilePath��  ��   m   � � �    - d   l  � ����� n   � � 1   � ��
� 
strq o   � ��~�~ &0 zipextractionpath zipExtractionPath��  ��  ��    l  � ��}�}   6 0 ...keep a local copy in the resources folder...    � `   . . . k e e p   a   l o c a l   c o p y   i n   t h e   r e s o u r c e s   f o l d e r . . .  I   � ��|�{�| 0 
copyfolder 
CopyFolder �z b   � � !  b   � �"#" b   � �$%$ b   � �&'& b   � �()( o   � ��y�y &0 zipextractionpath zipExtractionPath) m   � �** �++  / D i a l o g _ T o o l k i t' o   � ��x�x $0 dialogbundlename dialogBundleName% m   � �,, �--  - , -# o   � ��w�w "0 resourcesfolder resourcesFolder! o   � ��v�v $0 dialogbundlename dialogBundleName�z  �{   ./. l  � ��u01�u  0 ; 5 ...and copy the script bundle to the required folder   1 �22 j   . . . a n d   c o p y   t h e   s c r i p t   b u n d l e   t o   t h e   r e q u i r e d   f o l d e r/ 3�t3 I   � ��s4�r�s 0 
copyfolder 
CopyFolder4 5�q5 b   � �676 b   � �898 b   � �:;: b   � �<=< o   � ��p�p &0 zipextractionpath zipExtractionPath= m   � �>> �??  / D i a l o g _ T o o l k i t; o   � ��o�o $0 dialogbundlename dialogBundleName9 m   � �@@ �AA  - , -7 o   � ��n�n 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�q  �r  �t  � R      �m�l�k
�m .ascrerr ****      � ****�l  �k  ��  ��  ��  � BCB l  � ��j�i�h�j  �i  �h  C DED l  � ��gFG�g  F D > Remove unneeded files and folders created during this process   G �HH |   R e m o v e   u n n e e d e d   f i l e s   a n d   f o l d e r s   c r e a t e d   d u r i n g   t h i s   p r o c e s sE IJI I   � ��fK�e�f 0 
deletefile 
DeleteFileK L�dL o   � ��c�c 0 zipfilepath zipFilePath�d  �e  J MNM I   � ��bO�a�b 0 deletefolder DeleteFolderO P�`P o   � ��_�_ &0 zipextractionpath zipExtractionPath�`  �a  N QRQ l  � ��^�]�\�^  �]  �\  R STS l  � ��[UV�[  U V P One final check to verify installation was successful and return true if it was   V �WW �   O n e   f i n a l   c h e c k   t o   v e r i f y   i n s t a l l a t i o n   w a s   s u c c e s s f u l   a n d   r e t u r n   t r u e   i f   i t   w a sT X�ZX L   � �YY I   � ��YZ�X�Y "0 doesbundleexist DoesBundleExistZ [�W[ o   � ��V�V 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�W  �X  �Z  u \]\ l     �U�T�S�U  �T  �S  ] ^_^ i   p s`a` I      �Rb�Q�R 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlusb c�Pc o      �O�O "0 resourcesfolder resourcesFolder�P  �Q  a k     Udd efe r     ghg b     	iji n     klk 1    �N
�N 
psxpl l    m�M�Lm I    �Kn�J
�K .earsffdralis        afdrn m     �I
�I afdrcusr�J  �M  �L  j m    oo �pp d L i b r a r y / S c r i p t   L i b r a r i e s / D i a l o g   T o o l k i t   P l u s . s c p t dh o      �H�H 20 dialogtoolkitplusbundle dialogToolkitPlusBundlef qrq r    sts b    uvu o    �G�G "0 resourcesfolder resourcesFolderv m    ww �xx 4 / D i a l o g   T o o l k i t   P l u s . s c p t dt o      �F�F 0 	localcopy 	localCopyr yzy l   �E�D�C�E  �D  �C  z {|{ Z    R}~�B} I    �A��@�A "0 doesbundleexist DoesBundleExist� ��?� o    �>�> 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�?  �@  ~ Q    L���� k    A�� ��� Z   6���=�<� H    %�� I    $�;��:�; "0 doesbundleexist DoesBundleExist� ��9� o     �8�8 0 	localcopy 	localCopy�9  �:  � I   ( 2�7��6�7 0 
copyfolder 
CopyFolder� ��5� b   ) .��� b   ) ,��� o   ) *�4�4 20 dialogtoolkitplusbundle dialogToolkitPlusBundle� m   * +�� ���  - , -� o   , -�3�3 0 	localcopy 	localCopy�5  �6  �=  �<  � ��� I   7 =�2��1�2 0 deletefolder DeleteFolder� ��0� o   8 9�/�/ 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�0  �1  � ��.� r   > A��� m   > ?�-
�- boovtrue� o      �,�, 0 removalresult removalResult�.  � R      �+�*�)
�+ .ascrerr ****      � ****�*  �)  � r   I L��� m   I J�(
�( boovfals� o      �'�' 0 removalresult removalResult�B   r   O R��� m   O P�&
�& boovtrue� o      �%�% 0 removalresult removalResult| ��� l  S S�$�#�"�$  �#  �"  � ��!� L   S U�� o   S T� �  0 removalresult removalResult�!  _ ��� l     ����  �  �  �       ��������������������������������  � �����������������
�	��������� ����� 00 getscriptversionnumber GetScriptVersionNumber� "0 getmacosversion GetMacOSVersion� 0 splitstring SplitString� "0 loadapplication LoadApplication� 0 isapploaded IsAppLoaded� "0 closepowerpoint ClosePowerPoint� .0 changefilepermissions ChangeFilePermissions� $0 comparemd5hashes CompareMD5Hashes� 0 copyfile CopyFile� 00 createzipwithlocal7zip CreateZipWithLocal7Zip� <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiver� 0 
deletefile 
DeleteFile� "0 doesbundleexist DoesBundleExist� 0 doesfileexist DoesFileExist� 0 downloadfile DownloadFile� 0 findsignature FindSignature�
 0 installfonts InstallFonts�	 0 
renamefile 
RenameFile� 0 savepptaspdf SavePptAsPdf� 0 clearfolder ClearFolder� .0 clearpdfsafterzipping ClearPDFsAfterZipping� 0 
copyfolder 
CopyFolder� 0 createfolder CreateFolder� 0 deletefolder DeleteFolder� "0 doesfolderexist DoesFolderExist� 80 installdialogdisplayscript InstallDialogDisplayScript�  >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder�� 40 installdialogtoolkitplus InstallDialogToolkitPlus�� 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus� �� ���������� 00 getscriptversionnumber GetScriptVersionNumber�� ����� �  ���� 0 paramstring paramString��  � ���� 0 paramstring paramString� ���� 4���� �� �� &���������� "0 getmacosversion GetMacOSVersion�� ����� �  ���� 0 paramstring paramString��  � ������ 0 paramstring paramString�� 0 	osversion 	osVersion�  8������
�� .sysoexecTEXT���     TEXT��  ��  ��  �j E�O�W X  h� �� H���������� 0 splitstring SplitString�� ����� �  ������ &0 passedparamstring passedParamString�� (0 parameterseparator parameterSeparator��  � ���������� &0 passedparamstring passedParamString�� (0 parameterseparator parameterSeparator�� 00 oldtextitemsdelimiters oldTextItemsDelimiters�� *0 separatedparameters separatedParameters� ������
�� 
ascr
�� 
txdl
�� 
citm��  � *�,E�O�*�,FO��-E�O�*�,FUO�� �� w���������� "0 loadapplication LoadApplication�� ����� �  ���� 0 appname appName��  � �������� 0 appname appName�� 0 errmsg errMsg�� 0 errnum errNum� 	���� ���� ��� � �
�� 
capp
�� .miscactvnull��� ��� null�� 0 errmsg errMsg� ������
�� 
errn�� 0 errnum errNum��  
�� 
spac�� * *�/ *j UO�W X  ��%�%�%�%�%�%� �� ����������� 0 isapploaded IsAppLoaded�� ����� �  ���� 0 appname appName��  � ���������� 0 appname appName�� 0 
loadresult 
loadResult�� 0 errmsg errMsg�� 0 errnum errNum�  ������� � ���� � � �
�� 
prcs
�� 
pnam
�� 
spac�� 0 errmsg errMsg� ������
�� 
errn�� 0 errnum errNum��  �� ; (� *�-�,� ��%�%E�Y 	��%�%E�UO�W X  �%�%�%�%�%� �� ����������� "0 closepowerpoint ClosePowerPoint�� ����� �  ���� 0 paramstring paramString��  � ������ 0 paramstring paramString�� 0 closeresult closeResult� ����	������
�� 
prcs
�� 
pnam
�� .aevtquitnull��� ��� null��  ��  �� 4 +� #*�-�,� � *j UO�E�Y �E�O�UW 	X  	�� ��,���������� .0 changefilepermissions ChangeFilePermissions�� ����� �  ���� 0 paramstring paramString��  � ���������� 0 paramstring paramString��  0 newpermissions newPermissions�� 0 filepath filePath�� $0 quarantinestatus quarantineStatus� 8����S������\c����~�� 0 splitstring SplitString
�� 
cobj
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� g*��l+ E[�k/E�Z[�l/E�ZO (��%��,%j E�O�� ��%��,%j Y hW X 	 
hO ��%�%�%��,%j OeW 	X 	 
f� ������������� $0 comparemd5hashes CompareMD5Hashes�� ����� �  ���� 0 paramstring paramString��  � ���������� 0 paramstring paramString�� 0 filepath filePath�� 0 	validhash 	validHash�� 0 checkresult checkResult� 
�������������������� 0 splitstring SplitString
�� 
cobj�� 0 doesfileexist DoesFileExist
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� H*��l+ E[�k/E�Z[�l/E�ZO*�k+  fY hO ��%��,%j E�O�� W 	X  	f� ������������� 0 copyfile CopyFile�� ����� �  ���� 0 	filepaths 	filePaths��  � �������� 0 	filepaths 	filePaths�� 0 
targetfile 
targetFile�� "0 destinationfile destinationFile� 	����������~�}�|�� 0 splitstring SplitString
�� 
cobj
�� 
spac
� 
strq
�~ .sysoexecTEXT���     TEXT�}  �|  �� 9*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW 	X  f� �{�z�y���x�{ 00 createzipwithlocal7zip CreateZipWithLocal7Zip�z �w��w �  �v�v 0 
zipcommand 
zipCommand�y  � �u�t�u 0 
zipcommand 
zipCommand�t 0 errmsg errMsg� �s�r�q
�s .sysoexecTEXT���     TEXT�r  �q  �x  �j  O�W 	X  �� �p�o�n���m�p <0 createzipwithdefaultarchiver CreateZipWithDefaultArchiver�o �l��l �  �k�k 0 paramstring paramString�n  � �j�i�h�g�j 0 paramstring paramString�i 0 savepath savePath�h 0 zippath zipPath�g 0 errmsg errMsg� &�f�e@�d�cDH�bL�a�`�f 0 splitstring SplitString
�e 
cobj
�d 
spac
�c 
strq
�b .sysoexecTEXT���     TEXT�a  �`  �m =*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%�%�%j O�W 	X 
 �� �_T�^�]���\�_ 0 
deletefile 
DeleteFile�^ �[��[ �  �Z�Z 0 filepath filePath�]  � �Y�Y 0 filepath filePath� n�X�W�V�U�T
�X 
spac
�W 
strq
�V .sysoexecTEXT���     TEXT�U  �T  �\  ��%��,%j OeW 	X  f� �S{�R�Q���P�S "0 doesbundleexist DoesBundleExist�R �O��O �  �N�N 0 
bundlepath 
bundlePath�Q  � �M�M 0 
bundlepath 
bundlePath� ��L�K
�L 
ditm
�K .coredoexnull���     ****�P � *�/j U� �J��I�H���G�J 0 doesfileexist DoesFileExist�I �F��F �  �E�E 0 filepath filePath�H  � �D�D 0 filepath filePath� ��C�B�A�@�?
�C 
ditm
�B .coredoexnull���     ****
�A 
pcls
�@ 
file
�? 
bool�G � *�/j 	 *�/�,� �&U� �>��=�<���;�> 0 downloadfile DownloadFile�= �:��: �  �9�9 0 paramstring paramString�<  � �8�7�6�8 0 paramstring paramString�7 "0 destinationpath destinationPath�6 0 fileurl fileURL� ��5�4��3�2�1�0�/��.�5 0 splitstring SplitString
�4 
cobj
�3 
spac
�2 
strq
�1 .sysoexecTEXT���     TEXT�0  �/  
�. .sysodlogaskr        TEXT�; C*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW X  ��%�%j 
Of� �-��,�+���*�- 0 findsignature FindSignature�, �)��) �  �(�( 0 signaturepath signaturePath�+  � �'�' 0 signaturepath signaturePath� 	�&
�%�$�& 0 doesfileexist DoesFileExist�%  �$  �* 4 +*��%k+  	��%Y *��%k+  	��%Y �W 	X  �� �#%�"�!��� �# 0 installfonts InstallFonts�" ��� �  �� 0 paramstring paramString�!  � ����� 0 paramstring paramString� 0 fontname fontName� 0 fonturl fontURL� 0 fontpath fontPath� 
1�����C�b�� 0 splitstring SplitString
� 
cobj
� afdrcusr
� .earsffdralis        afdr
� 
psxp� 0 doesfileexist DoesFileExist� 0 downloadfile DownloadFile�  A*��l+ E[�k/E�Z[�l/E�ZO�j �,�%�%E�O*�k+  eY hO*��%�%k+ 	� �i������ 0 
renamefile 
RenameFile� ��� �  �� 0 paramstring paramString�  � ���
� 0 paramstring paramString� 0 
targetfile 
targetFile�
 0 newfilename newFilename� 
z�	���������	 0 splitstring SplitString
� 
cobj
� 
psxp
� 
strq
� 
spac
� .sysoexecTEXT���     TEXT�  �  � E*��l+ E[�k/E�Z[�l/E�ZO��,�,E�O��,�,E�O ��%�%�%�%j OeW 	X  	f� ��� ������� 0 savepptaspdf SavePptAsPdf�  ����� �  ���� 0 tempsavepath tempSavePath��  � ������ 0 tempsavepath tempSavePath�� 0 thisdocument thisDocument� 
�������������������
�� 
AAPr
�� 
kfil
�� 
psxf
�� 
fltp
�� pSAT � �� 
�� .coresavenull���     obj ��  ��  �� ( � *�,E�O��*�/��� UOeW 	X  	f� ������������� 0 clearfolder ClearFolder�� �� ��    ���� 0 foldertoempty folderToEmpty��  � ���� 0 foldertoempty folderToEmpty� �������������
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� @ 7��%��,%�%�%j O��%��,%�%�%j O��%��,%�%�%j OeW 	X 	 
f� ���������� .0 clearpdfsafterzipping ClearPDFsAfterZipping�� ����   ���� 0 foldertoempty folderToEmpty��   ���� 0 foldertoempty folderToEmpty 0����5������
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  ��   ��%��,%�%�%j OeW 	X  f� ��?�������� 0 
copyfolder 
CopyFolder�� ����   ���� 0 
folderpath 
folderPath��   �������� 0 
folderpath 
folderPath�� 0 targetfolder targetFolder�� &0 destinationfolder destinationFolder 	P����f������������ 0 splitstring SplitString
�� 
cobj
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  �� 9*��l+ E[�k/E�Z[�l/E�ZO ��%��,%�%��,%j OeW 	X  f� ��v�������� 0 createfolder CreateFolder�� ��	�� 	  ���� 0 
folderpath 
folderPath��   ���� 0 
folderpath 
folderPath �����������
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  ��  ��%��,%j OeW 	X  f� �������
���� 0 deletefolder DeleteFolder�� ����   ���� 0 
folderpath 
folderPath��  
 ���� 0 
folderpath 
folderPath �����������
�� 
spac
�� 
strq
�� .sysoexecTEXT���     TEXT��  ��  ��  ��%��,%j OeW 	X  f� ����������� "0 doesfolderexist DoesFolderExist�� ����   ���� 0 
folderpath 
folderPath��   ���� 0 
folderpath 
folderPath �����������
�� 
ditm
�� .coredoexnull���     ****
�� 
pcls
�� 
cfol
�� 
bool�� � *�/j 	 *�/�,� �&U� ����������� 80 installdialogdisplayscript InstallDialogDisplayScript�� ����   ���� 0 paramstring paramString��   �������� 0 paramstring paramString�� 0 
scriptpath 
scriptPath�� 0 downloadurl downloadURL ����������
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp�� 0 downloadfile DownloadFile�� �j �,�%E�O�E�O*��%�%k+ � ���������� >0 checkforscriptlibrariesfolder CheckForScriptLibrariesFolder�� ����   ���� 0 paramstring paramString��   ������ 0 paramstring paramString�� .0 scriptlibrariesfolder scriptLibrariesFolder ������"��>��������R����h����p
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
sisn��  ��  �� ^�j �,�%E�O*�k+  �Y E 9��%��,%�el 	O�*j �,�,%�%��,%�el 	O���,%�el 	O�W X  a � ��w�������� 40 installdialogtoolkitplus InstallDialogToolkitPlus�� ��   �~�~ "0 resourcesfolder resourcesFolder��   �}�|�{�z�y�x�w�} "0 resourcesfolder resourcesFolder�| 0 downloadurl downloadURL�{ .0 scriptlibrariesfolder scriptLibrariesFolder�z $0 dialogbundlename dialogBundleName�y 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�x 0 zipfilepath zipFilePath�w &0 zipextractionpath zipExtractionPath �v�u�t�����s�r�q�p�o��n��m�l�k�j*,>@�i�h
�v afdrcusr
�u .earsffdralis        afdr
�t 
psxp�s "0 doesbundleexist DoesBundleExist�r "0 doesfolderexist DoesFolderExist�q 0 createfolder CreateFolder�p  �o  �n 0 
copyfolder 
CopyFolder�m 0 downloadfile DownloadFile
�l 
spac
�k 
strq
�j .sysoexecTEXT���     TEXT�i 0 
deletefile 
DeleteFile�h 0 deletefolder DeleteFolder�� ��E�O�j �,�%E�O�E�O��%E�O��%E�O��%E�O*�k+  eY hO*�k+ 	  *�k+ 
W 	X  fY hO*��%k+  *��%�%�%k+  eY hY hO*��%�%k+  T Ha _ %�a ,%a %�a ,%j O*�a %�%a %�%�%k+ O*�a %�%a %�%k+ W X  hY hO*�k+ O*�k+ O*�k+ � �ga�f�e�d�g 80 uninstalldialogtoolkitplus UninstallDialogToolkitPlus�f �c�c   �b�b "0 resourcesfolder resourcesFolder�e   �a�`�_�^�a "0 resourcesfolder resourcesFolder�` 20 dialogtoolkitplusbundle dialogToolkitPlusBundle�_ 0 	localcopy 	localCopy�^ 0 removalresult removalResult �]�\�[ow�Z��Y�X�W�V
�] afdrcusr
�\ .earsffdralis        afdr
�[ 
psxp�Z "0 doesbundleexist DoesBundleExist�Y 0 
copyfolder 
CopyFolder�X 0 deletefolder DeleteFolder�W  �V  �d V�j �,�%E�O��%E�O*�k+  6 (*�k+  *��%�%k+ Y hO*�k+ OeE�W 
X 	 
fE�Y eE�O� ascr  ��ޭ