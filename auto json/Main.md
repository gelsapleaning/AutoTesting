〈〈 JSONファイル生成ツール 〉〉								データ追加
								

■前提												
    当ツールは、「backendService　API　Response」のJSONファイル作成ツールになります。												
												
■使い方（各種定義シート）												
    1.　API レスポンスJSONのレイアウト定義をシートとして追加してください。												
    2.　データ一覧の最下行より下には一切何も記載しないでください。　（→　なにかあればそのセルまで含めて一覧の対象と見なします。）												
    3.　配列項目（Array定義）の編集について												
　　　　1）Array項目については、以下のように記載を変更してください。												
              'A.　子項目レベルでManyではなく、その一つ上の階層（親項目レベル）でManyを指定して下さい。												
	（例）（各種API定義書）						（当定義シート上）					
	Item Name				Array		Item Name				Array	
	list						list				Many	
	referenceNo				Many		referenceNo					
	beneficiaryName				Many		beneficiaryName					
	beneficiarybranch				Many		beneficiarybranch					
	beneficiaryBank				Many		beneficiaryBank					
	transactionDate				Many		transactionDate					
	batchNo				Many		batchNo					
	beneficaryAccountNo				Many		beneficaryAccountNo				Break	
												
             'B.　子項目で一つのリストグループ化を判別する為、グループ化の末端項目にBreakを指定して下さい。												
												
             C.　リストグループを複数定義したい場合は、以下のように行を複製して下さい。												
	（例）（各種API定義書）						（当定義シート上）					
	Item Name				Array		Item Name				Array	
	list						list				Many	
	referenceNo				Many		referenceNo					
	beneficiaryName				Many		beneficiaryName					
	beneficiarybranch				Many		beneficiarybranch					
	beneficiaryBank				Many		beneficiaryBank					
	transactionDate				Many		transactionDate					
	batchNo				Many		batchNo					
	beneficaryAccountNo				Many		beneficaryAccountNo				Break	←　1グループの末端として"Break"を指定
							referenceNo					
							beneficiaryName					
							beneficiarybranch					
							beneficiaryBank					
							transactionDate					
							batchNo					
							beneficaryAccountNo				Break	←　2グループの末端として"Break"を指定
												
             D.　複数データを作成する際に、リストの数を他の列のデータより少なくしたい場合は、 データを設定しないすべてのValue欄に #noItems を指定してください。(03_Mutual Fund Account Overviewシート参照)												
	※きっちりリスト一つ分に指定しないとJSON形式が崩れる可能性があります。											
												
■使い方												
    1.　「①シート読込」ボタンを押下して、ファイル作成対象シートをリスト化します。												
    2.　ファイル作成対象シートを選択して、「JSON作成」ボタンを押下してください。												
         「ALL」を選択すると、データが入っているすべてのシートのJSONを作成します。												
