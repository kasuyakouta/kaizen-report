\# 改善報告アプリ - プロジェクト概要



\## 概要

産廃回収 改善報告管理アプリ。

URL: https://kasuyakouta.github.io/kaizen-report/



\## 技術スタック

\- フロント: GitHub Pages上の単一HTMLファイル

\- バックエンド: Google Apps Script (GAS)

\- GAS URL: https://script.google.com/macros/s/AKfycbw3x\_oa9bvka9SlpuKDZ5U1YU7xKVMLY1u6QLpLv1NXXIalIYfR1p0Dxappb3E6WvdK/exec

\- データ保存: Google Sheets



\## このアプリ特有の設計(他アプリと異なる点に注意)

\- \*\*GAS通信は隠しiframe+フォーム送信方式\*\*(gasPost関数)。fetchのtext/plain方式ではない。CORS・GASのリダイレクト問題を根本的に回避するための実装のため、この方式は変更しないこと

\- \*\*管理者PINは共通の3150ではなく、このアプリ内で個別に設定・変更できる4桁PIN\*\*(SHA-256ハッシュ化してlocalStorageに保存)

\- Service Workerは未実装(manifest.jsonのみ)。オフライン対応は現状ない

\- 月次目標(monthlyGoal)・年間目標(annualGoal)の設定機能あり



\## 必須ルール(標準スタック)

\- iOS Safari互換性を優先の設計制約とする

\- 日時はローカル時刻で組み立てる(UTCは使わない)

\- フォントは IBM Plex Sans JP / Noto Sans JP(未使用の場合は既存フォントを維持)



\## 変更時のお願い

\- 複雑な変更は実装前にオプションA/B形式で提案し、承認を得てから実装する

\- 回答は簡潔に、前置きは省略する

\- 上記の「このアプリ特有の設計」は、明確な指示がない限り変更しないこと

