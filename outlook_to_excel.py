import os
import win32com.client
import pandas as pd
from datetime import datetime

# ==========================================
# 1. 基本設定
# ==========================================

# 集計対象の期間を指定（YYYY/MM/DD形式）
START_DATE = "⚪︎⚪︎⚪︎⚪︎/⚪︎⚪︎/⚪︎⚪︎"
END_DATE = "⚪︎⚪︎⚪︎⚪︎/⚪︎⚪︎/⚪︎⚪︎"

# 出力先Excelファイルのパス設定
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
excel_folder_path = os.path.join(desktop_path, "data")
EXCEL_FILE_PATH = os.path.join(
    excel_folder_path, "sample.xlsx"
)

# 検索対象のOutlookアカウントとフォルダ構成
TARGET_EMAIL_ACCOUNT = "sample@sample.co.jp"
TARGET_FOLDER_NAME = "受信トレイ"
TARGET_SUBFOLDER_NAME = "任意のフォルダ名"

# Excel出力時の列順定義
COLUMN_ORDER = [
    "受信時刻",
    "種別",
    "名前",
    "メールアドレス",
]

# メールの件名末尾と「種別」の対応マッピング
SUBJECT_MAPPING = {
    "〇〇のお問い合わせがありました": "種別1",
    "資料請求のお問い合わせ": "種別2",
}


def parse_email_body(body):
    """
    メール本文から「■」または「▼」で始まる項目名と、その次の行の値を抽出。
    """
    data = {}
    lines = body.splitlines()
    for i, line in enumerate(lines):
        stripped_line = line.strip()
        if stripped_line.startswith(("■", "▼")):
            # 先頭の記号を除去してキー名を取得
            key = stripped_line[1:].strip()
            value = ""
            # 次の項目が始まるまでの間の非空行を値として取得
            for j in range(i + 1, len(lines)):
                next_line = lines[j].strip()
                if next_line.startswith(("■", "▼")):
                    break
                if next_line:
                    value = next_line
                    break
            data[key] = value
    return data


def main():
    print("処理を開始します...")
    os.makedirs(excel_folder_path, exist_ok=True)

    # ==========================================
    # 2. Outlookへの接続と対象フォルダの取得
    # ==========================================
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        print(f"Outlookへの接続に失敗しました: {e}")
        return

    target_folder = None
    try:
        account = outlook.Folders[TARGET_EMAIL_ACCOUNT]
        inbox = account.Folders[TARGET_FOLDER_NAME]
        target_folder = inbox.Folders[TARGET_SUBFOLDER_NAME]
    except Exception as e:
        print(f"フォルダの取得に失敗しました: {e}")
        return

    print(
        f"'{TARGET_EMAIL_ACCOUNT}' の '{target_folder.Name}' フォルダをスキャンします。"
    )
    print(f"集計期間: {START_DATE} から {END_DATE}")

    # ==========================================
    # 3. 既存Excelファイルの読み込みと重複チェック用データの作成
    # ==========================================
    df_existing = pd.DataFrame(columns=COLUMN_ORDER)
    if os.path.exists(EXCEL_FILE_PATH):
        try:
            print(f"既存のExcelファイル '{EXCEL_FILE_PATH}' を読み込みます。")
            df_existing = pd.read_excel(EXCEL_FILE_PATH)
        except Exception as e:
            print(f"Excelファイルの読み込みに失敗しました: {e}")
            return
    else:
        print("Excelファイルが存在しないため、新規に作成します。")

    # 既存データから (名前, メールアドレス, 種別) の組み合わせをセットに登録し、重複排除に利用
    seen_combinations = set()
    if not df_existing.empty:
        print("既存データの重複チェック用セットを作成します...")

        required_columns = ["名前", "メールアドレス", "種別"]
        if all(col in df_existing.columns for col in required_columns):
            for index, row in df_existing.iterrows():
                name = "" if pd.isna(row["名前"]) else str(row["名前"])
                email = (
                    "" if pd.isna(row["メールアドレス"]) else str(row["メールアドレス"])
                )
                inquiry_type = "" if pd.isna(row["種別"]) else str(row["種別"])

                # 名前またはメールアドレスが存在する場合にのみ登録
                if name or email:
                    seen_combinations.add((name, email, inquiry_type))
        print(f"既存の重複チェック用キーを {len(seen_combinations)} 件作成しました。")

    # ==========================================
    # 4. 対象期間のメール抽出とデータ解析
    # ==========================================
    start_date_str = datetime.strptime(START_DATE, "%Y/%m/%d").strftime(
        "%Y-%m-%d %H:%M"
    )
    end_date_str = datetime.strptime(END_DATE + " 23:59", "%Y/%m/%d %H:%M").strftime(
        "%Y-%m-%d %H:%M"
    )
    date_filter = f"([ReceivedTime] >= '{start_date_str}') AND ([ReceivedTime] <= '{end_date_str}')"

    # 指定期間内の通常メール（IPM.Note）のみを抽出
    filter_str = f"([MessageClass] = 'IPM.Note') AND {date_filter}"
    messages = target_folder.Items.Restrict(filter_str)
    messages.Sort("[ReceivedTime]", True)

    print(f"{len(messages)} 件のメールをチェックします...")

    new_records = []
    for message in messages:
        try:
            # テストメールは除外
            if "テスト" in message.Subject:
                continue

            # 件名から該当する種別を判定
            matched_key = None
            for key in SUBJECT_MAPPING.keys():
                if message.Subject.strip().endswith(key):
                    matched_key = key
                    break

            if matched_key:
                current_inquiry_type = SUBJECT_MAPPING[matched_key]

                email_data = {
                    "Subject": message.Subject,
                    "Body": message.Body,
                    "ReceivedTime": message.ReceivedTime,
                }

                # タイムゾーン情報を除去
                received_time_naive = email_data["ReceivedTime"].replace(tzinfo=None)
                body_data = parse_email_body(email_data["Body"])
                email_address = body_data.get("メールアドレス", "")

                # 件名から「名前」を抽出（例: "〇〇様から..."）
                name_from_subject = ""
                if "様から" in email_data["Subject"]:
                    name_from_subject = email_data["Subject"].split("様から")[0].strip()

                # 重複判定: (名前, メールアドレス, 種別) が既に存在するか確認
                is_duplicate = False
                if name_from_subject or email_address:
                    key = (name_from_subject, email_address, current_inquiry_type)

                    if key in seen_combinations:
                        is_duplicate = True
                    else:
                        seen_combinations.add(key)

                if is_duplicate:
                    continue

                # 新規レコードの作成
                record = {
                    "受信時刻": received_time_naive,
                    "種別": current_inquiry_type,
                    **body_data,
                    "名前": name_from_subject,
                }

                new_records.append(record)
                print(f"新規データを追加: {email_data['Subject']}")

        except Exception as e:
            print(f"--- エラーが発生したため、一件のメールをスキップしました ---")
            try:
                print(f"対象メール件名: {message.Subject}")
            except:
                print("対象メールの件名取得にも失敗しました。")
            print(f"エラー内容: {e}")
            continue

    # ==========================================
    # 5. 抽出データのExcelへの追記保存
    # ==========================================
    if new_records:
        df_new = pd.DataFrame(new_records)

        # 既存データと新規データを結合
        if df_existing.empty:
            df_combined = df_new
        else:
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)

        # 受信時刻の降順（新しい順）でソートし、列順を整理
        df_sorted = df_combined.sort_values(by="受信時刻", ascending=False)
        df_final = df_sorted.reindex(columns=COLUMN_ORDER)

        try:
            df_final.to_excel(EXCEL_FILE_PATH, index=False, engine="openpyxl")
            print(
                f"\n処理が完了しました。{len(new_records)} 件の新規データを追記し、'{EXCEL_FILE_PATH}' を更新しました。"
            )
        except Exception as e:
            print(f"\nExcelファイルへの保存中にエラーが発生しました: {e}")
    else:
        print("\n追記する新しいメールはありませんでした。")


if __name__ == "__main__":
    main()
