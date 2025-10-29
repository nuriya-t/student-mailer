

import os, sys, time, json
import pandas as pd
import msal, requests

# ====== КОНФИГ ======
CLIENT_ID = "4e7eafee-f93e-4e78-a003-1e5f3a0835b3"
TENANT_ID = "organizations"
SCOPES = ["Mail.Send"]

EXCEL_PATH = r"C:\Users\nuriy\Downloads\student_debt.xlsx"
GRAPH_SEND_URL = "https://graph.microsoft.com/v1.0/me/sendMail"

SEND_DELAY_SEC = 0.4
DRY_RUN = False
# =====================


def ensure_columns(df):
    need = {"Email", "ФИО", "Дисциплина", "Факультет", "Уровень"}
    miss = need - set(df.columns)
    if miss:
        raise ValueError("В Excel нет столбцов: " + ", ".join(miss))


def pick_recipients(df):
    print("\n📩 Кому хотите отправить письма?")
    print("1 - Всем студентам")
    print("2 - Определённому студенту по email")
    print("3 - Всем студентам по дисциплине")
    print("4 - Всем студентам по уровню (Бакалавр / Магистратура)")
    choice = input("Введите номер варианта: ").strip()

    df["Email"] = df["Email"].astype(str).str.strip()
    df["Дисциплина"] = df["Дисциплина"].astype(str).str.strip()
    df["Уровень"] = df["Уровень"].astype(str).str.strip()

    if choice == "1":
        recips = df["Email"].dropna().tolist()
    elif choice == "2":
        email = input("Введите email студента: ").strip()
        recips = df.loc[df["Email"].str.lower() == email.lower(), "Email"].tolist()
    elif choice == "3":
        disc = input("Введите название дисциплины: ").strip()
        recips = df.loc[df["Дисциплина"].str.lower() == disc.lower(), "Email"].tolist()
    elif choice == "4":
        level = input("Введите уровень (Бакалавр / Магистратура): ").strip().lower()
        recips = df.loc[df["Уровень"].str.lower() == level, "Email"].tolist()
    else:
        recips = []

    recips = sorted({r for r in recips if r and r != "nan"})
    return recips


def build_subject_body(student_name, discipline, faculty):
    subject = f"Задолженность по дисциплине {discipline}"
    html = f"""
    <html>
    <body style="font-family: Arial; color: #333;">
      <p>Уважаемый(ая) <b>{student_name}</b>,</p>
      <p>Сообщаем Вам, что у Вас имеется задолженность по дисциплине <b>"{discipline}"</b>.</p>
      <p>Просим в кратчайшие сроки связаться с преподавателем или учебным офисом для её устранения.</p>
      <p style="margin-top:20px;">
         С уважением,<br>
         <b>Деканат факультета {faculty}</b><br>
         Астана IT Университет
      </p>
    </body>
    </html>
    """.strip()
    return subject, html


def acquire_token():
    """Авторизация через MSAL Device Code Flow"""
    tried = []

    def _try_authority(auth_tenant):
        authority = f"https://login.microsoftonline.com/{auth_tenant}"
        app = msal.PublicClientApplication(CLIENT_ID, authority=authority)
        result = app.acquire_token_silent(SCOPES, account=None)
        if result and "access_token" in result:
            return result["access_token"]

        try:
            flow = app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                raise RuntimeError(flow.get("error_description") or "Device Code Flow init failed")
            print(f"\n🔐 Откройте {flow['verification_uri']} и введите код: {flow['user_code']}\n")
            result = app.acquire_token_by_device_flow(flow)
            if "access_token" in result:
                return result["access_token"]
            raise RuntimeError(result.get("error_description") or str(result))
        except Exception as e:
            tried.append((authority, f"device_code: {e}"))
            return None

    tok = _try_authority(TENANT_ID)
    if tok:
        return tok

    tok = _try_authority("organizations")
    if tok:
        return tok

    tried.append(("interactive/organizations", "Не удалось инициализировать вход"))
    msg = ["Не удалось получить токен (подробности):"]
    for a, err in tried:
        msg.append(f" - {a}: {err}")
    raise RuntimeError("\n".join(msg))


def send_mail_graph(access_token, to_email, subject, html_body):
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
        },
        "saveToSentItems": True
    }

    if DRY_RUN:
        print(f"[DRY-RUN] → {to_email} | {subject}")
        return True, None

    resp = requests.post(
        GRAPH_SEND_URL,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        },
        data=json.dumps(payload),
        timeout=30
    )
    if resp.status_code in (200, 202):
        return True, None
    else:
        try:
            detail = resp.json()
        except Exception:
            detail = resp.text
        return False, f"{resp.status_code}: {detail}"


def main():
    if not os.path.exists(EXCEL_PATH):
        print(f"❌ Файл не найден: {EXCEL_PATH}")
        sys.exit(1)

    df = pd.read_excel(EXCEL_PATH)
    ensure_columns(df)


    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    print("\n📊 Данные из Excel:\n")
    print(df[["ФИО", "Email", "Дисциплина", "Факультет", "Уровень"]].to_string(index=False))
    print("\nВсего записей:", len(df))
   
    recipients = pick_recipients(df)
    if not recipients:
        print("⚠️ Не найдено получателей.")
        sys.exit(0)

    preview_df = df[df["Email"].isin(recipients)][["ФИО", "Email", "Дисциплина", "Факультет", "Уровень"]]
    print("\n📋 Список получателей:\n")
    print(preview_df.to_string(index=False))
    print(f"\nНайдено адресов: {len(recipients)}")

    if input("Продолжить отправку? (y/n): ").strip().lower() != "y":
        print("Отменено.")
        sys.exit(0)

    df_idx = df.set_index(df["Email"].str.lower())

    try:
        token = acquire_token()
    except Exception as e:
        print("❌ Ошибка авторизации:", e)
        sys.exit(1)

    ok = fail = 0
    for email in recipients:
        row = df_idx.loc[email.lower()]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]

        subject, html = build_subject_body(
            str(row["ФИО"]).strip(),
            str(row["Дисциплина"]).strip(),
            str(row["Факультет"]).strip()
        )

        success, err = send_mail_graph(token, email, subject, html)
        if success:
            print(f"✅ Отправлено: {email}")
            ok += 1
        else:
            print(f"❌ Ошибка для {email}: {err}")
            fail += 1

        time.sleep(SEND_DELAY_SEC)

    print("\n===== ИТОГ =====")
    print(f"Успешно: {ok}")
    print(f"Ошибки:  {fail}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nОстановлено пользователем.")
