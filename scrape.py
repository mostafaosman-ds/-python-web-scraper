from bs4 import BeautifulSoup
import pandas as pd
import requests


def scrape_quotes():
  # موقع تجريبي مخصص للتدريب وآمن تماماً من الحظر
  url = "http://quotes.toscrape.com/"
  print(f"[*] جاري الاتصال بالموقع التجريبي...")

  try:
    response = requests.get(url, timeout=15)

    if response.status_code == 200:
      print("[+] تم الاتصال بنجاح! جاري تحليل البيانات...")
      soup = BeautifulSoup(response.content, "html.parser")

      quotes_data = []
      # سحب الأقوال والمقولات كمثال عملي على الـ Scraping
      items = soup.find_all("div", class_="quote")

      print(f"[+] تم العثور على {len(items)} عنصر.")

      for item in items:
        text = item.find("span", class_="text").text.strip()
        author = item.find("small", class_="author").text.strip()

        quotes_data.append({"المقولة": text, "الكاتب": author})

      # تنظيف وحفظ البيانات في ملف Excel باستخدام Pandas
      if quotes_data:
        df = pd.DataFrame(quotes_data)
        df.drop_duplicates(inplace=True)

        print("\n--- عينة من البيانات المسحوبة ---")
        print(df.head(5))

        file_name = "quotes_results.xlsx"
        df.to_excel(file_name, index=False, engine="openpyxl")
        print(f"\n[تم بنجاح] تم حفظ البيانات في ملف Excel باسم: {file_name}")

        return df
      else:
        print("[-] لم يتم العثور على بيانات.")
    else:
      print(f"[-] فشل الاتصال، كود الاستجابة: {response.status_code}")

  except Exception as e:
    print(f"[-] حدث خطأ: {e}")


if __name__ == "__main__":
  scrape_quotes()