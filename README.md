# FortiGate Parameter Sheet Exporter

FortiGate の設定を REST API の GET で取得し、Excel パラメータシートとサニタイズ済み JSON へ出力する Python スクリプトです。

## 主な機能

- REST API は GET のみ使用
- 18 エンドポイントから設定を取得
- 21 枚の Excel シートを生成
- Firewall Policy が参照する Address、VIP、Service などを実値へ展開
- API 取得結果と参照解決結果を Validation シートへ記録
- API トークンは対話入力し、ファイルへ保存しない
- API の生レスポンスは保存しない
- シリアル番号と既定ホスト名をマスク

## 検証環境

- FortiGate-VM64
- FortiOS 8.0.0 build 167
- Ubuntu Desktop 26.04 LTS
- Python 3.14.4
- openpyxl 3.1.5

FortiOS のバージョンや Administrator Profile の権限によって、取得結果が異なる可能性があります。

## インストール

Ubuntu では openpyxl をインストールします。

```bash
sudo apt install -y python3-openpyxl
```

## 実行方法

```bash
python3 fortigate_parameter_export_v1_0.py \
  --base-url https://192.168.10.254 \
  --vdom root \
  --insecure
```

実行すると、FortiGate の API トークンの入力を求められます。トークンは画面へ表示されず、ファイルや環境変数にも保存されません。

`--insecure` は、自己署名証明書を使用する検証環境向けのオプションです。本番環境では信頼できる証明書を使用し、このオプションを付けずに実行してください。

## 出力ファイル

- `fortigate_parameter_sheet.xlsx`
- `fortigate_sanitized_snapshot.json`

出力ファイルは、実行日時を含むディレクトリへ保存されます。

## 注意事項

- このスクリプトは FortiGate のすべての設定を取得するものではありません。
- 対応範囲は、スクリプトに実装された 18 エンドポイントです。
- FortiOS のバージョンや Administrator Profile の権限によって、取得結果が異なる可能性があります。
- 現在のルーティングテーブルやリンク状態などの稼働ステータスは対象外です。
- Excel から FortiGate へ設定を戻すインポート機能は実装していません。
- `--insecure` は、自己署名証明書を使用する検証環境だけで使用してください。
- 本番環境で使用する前に、検証環境で動作を確認してください。
- Administrator Profile の権限は、必要な参照権限へ限定することを推奨します。
- Trusted Hosts は、スクリプトを実行するホストの IP アドレスへ限定することを推奨します。

## License

MIT License
