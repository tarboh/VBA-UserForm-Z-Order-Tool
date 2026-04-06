他言語のGUIのビジュアルエディタとしてVBAのユーザーフォームのエディタを活用する取り組みの中で、コントロールのZオーダー（重なり）を調べる需要があったため作ってみました。

 技術的な仕組み
このツールは、VBA標準の Controls コレクションに頼らず、ユーザーフォームの背後にある COM (Component Object Model) のインターフェース を直接操作することで、描画エンジンが管理している「真の重なり順」を抽出しています。

IOleContainer の活用:
ユーザーフォーム（MSForms.UserForm）を IOleContainer インターフェースとして扱い、EnumObjects メソッドを呼び出しています。ここから得られる列挙子（IEnumUnknown）の順序こそが、MSFormsが内部で管理している物理的なZオーダーそのものです。
再帰的な階層スキャン:Frame コレクションなどのコンテナ内にネストされたコントロールも網羅するため、再帰呼び出し構造を採用しています。これにより、複雑なレイアウトでも階層構造を保ったままスキャン可能です。
IUnknown によるポインタの正規化:
VBAの各コントロールの ObjPtr は、内部的なラッパーの影響でアドレスが変動する場合があります。本ツールでは、すべてのオブジェクトを IID_IUnknown で QueryInterface してアドレスを正規化（Identityの特定）してから照合しているため、100%の精度でコントロール名を特定できます。

 動作環境
32bit / 64bit 両対応（DispCallFunc による動的関数呼び出しを使用）
Windows環境のExcel/VBA

VBAのビジュアルエディタを高度なUI設計ツールとして活用する際の、バックエンド解析用パーツとして役立てていただければ幸いです。

使用上の注意点
オブジェクトに対してQueryInterfaceを実行する関係上、対象フォームのインスタンスが強制的に生成されます。
この影響で、フォームは表示されないけれどUserForm_Initializeイベントが実行されるという、VBAが本来想定していない挙動が発生します。
予期せぬクラッシュやフリーズを招く恐れがあるため、取り扱いには十分ご注意ください。 


<img width="545" height="427" alt="image" src="https://github.com/user-attachments/assets/23db8f86-43ad-43b8-97de-bd4926c9bd74" />
<img width="578" height="225" alt="image" src="https://github.com/user-attachments/assets/6dbd0862-401a-4834-8a59-fd301ee152f0" />
