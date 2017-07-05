01_Login Validate API												
Login	Validate											
Level	Item Name	description	Mandatory	Data type	Array	Min Length	Max Length	Format	Allowable Strings	Default Value	Sample Value	JP Note
L1	header											
L2	referenceNo	Unique reference number for each execution	Yes	String	-	20	30	-	[A-Z0-9]+	-	SMP400124788812345678	処理を一意に判別するための参照番号
L2	systemCode	Identifier of the Channel which invokes the requests	Yes	String	-	2	5	-	[A-Z]+	-	SMP	処理要求元を判別するためのシステムコード
L2	langCode	Language selected for the transaction	Yes	String	-	3	3	-	[A-Z]+	ENG/JAP	JAP	"【言語コード】
JAP - 日本語"
L1	requestParam											
L2	nationalid	PID of the Customer	Yes	String	-	10	10	-	[0-9]+	-	4001247888	個人識別番号（店番号＋口座番号）
L1	errorInfo											
L2	statusID	Error Code	Yes	String	-	5	5	-	[0-9]+	-	0	エラーコード（API Error Code参照）
L2	statusMessage	Error Message	Yes	String	-	7	200	-	[A-Z0-9]+	-	Success	エラーメッセージ（API Error Code参照）
L1	responseParam											
L2	status	Status of Login Transaction	Yes	String	-	7	200	-	[A-Z0-9]+	-	Success	"【ログイン実行結果】
Success - 成功"
L2	lastLoginTime	Last Login Time	Yes	String	-	21	21	YYYY/MM/DD HH24:MI:SS	[A-Z0-9:]+		2016/06/11 14:14:12 PM	前回ログイン日時
L2	flag	flag to indicate if Force Password screen has to be displayed	Yes	String	-	1	1	-	[0-9]	-	0 - To redirect to Force Change Password module	"【フラグ】
0 - 強制パスワード変更画面に遷移する"
