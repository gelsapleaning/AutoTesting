01_Login Validate API												
Login	Validate											
Level	Item Name	description	Mandatory	Data type	Array	Min Length	Max Length	Format	Allowable Strings	Default Value	Sample Value	JP Note
L1	header											
L2	referenceNo	Unique reference number for each execution	Yes	String	-	20	30	-	[A-Z0-9]+	-	SMP400124788812345678	��������ӂɔ��ʂ��邽�߂̎Q�Ɣԍ�
L2	systemCode	Identifier of the Channel which invokes the requests	Yes	String	-	2	5	-	[A-Z]+	-	SMP	�����v�����𔻕ʂ��邽�߂̃V�X�e���R�[�h
L2	langCode	Language selected for the transaction	Yes	String	-	3	3	-	[A-Z]+	ENG/JAP	JAP	"�y����R�[�h�z
JAP - ���{��"
L1	requestParam											
L2	nationalid	PID of the Customer	Yes	String	-	10	10	-	[0-9]+	-	4001247888	�l���ʔԍ��i�X�ԍ��{�����ԍ��j
L1	errorInfo											
L2	statusID	Error Code	Yes	String	-	5	5	-	[0-9]+	-	0	�G���[�R�[�h�iAPI Error Code�Q�Ɓj
L2	statusMessage	Error Message	Yes	String	-	7	200	-	[A-Z0-9]+	-	Success	�G���[���b�Z�[�W�iAPI Error Code�Q�Ɓj
L1	responseParam											
L2	status	Status of Login Transaction	Yes	String	-	7	200	-	[A-Z0-9]+	-	Success	"�y���O�C�����s���ʁz
Success - ����"
L2	lastLoginTime	Last Login Time	Yes	String	-	21	21	YYYY/MM/DD HH24:MI:SS	[A-Z0-9:]+		2016/06/11 14:14:12 PM	�O�񃍃O�C������
L2	flag	flag to indicate if Force Password screen has to be displayed	Yes	String	-	1	1	-	[0-9]	-	0 - To redirect to Force Change Password module	"�y�t���O�z
0 - �����p�X���[�h�ύX��ʂɑJ�ڂ���"
