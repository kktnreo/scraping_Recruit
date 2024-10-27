from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc
import time
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import sys
import os

# .envファイルを読み込む
load_dotenv()


def search_Recruitment(menu_num, work_place, income, free_word):
	if menu_num == 1:
		occupation = '営業職'
	elif menu_num == 2:
		occupation = '企画・管理'
	elif menu_num == 3:
		occupation = '事務・アシスタント'
	elif menu_num == 4:
		occupation = '販売・サービス職'
	elif menu_num == 5:
		occupation = '専門職（コンサルティングファーム・専門事務所・監査法人）'
	elif menu_num == 6:
		occupation = '金融系専門職'
	elif menu_num == 7:
		occupation = '公務員・教員・農林水産関連職'
	elif menu_num == 8:
		occupation = '技術職（SE・インフラエンジニア・Webエンジニア）'
	elif menu_num == 9:
		occupation = '技術職（機械・電気）'
	elif menu_num == 10:
		occupation = '技術職（組み込みソフトウェア）'
	elif menu_num == 11:
		occupation = '技術職・専門職（建設・建築・不動産・プラント・工場）'
	elif menu_num == 12:
		occupation = '技術職（化学・素材・化粧品・トイレタリー）'
	elif menu_num == 13:
		occupation = '技術職（食品・香料・飼料）'
	elif menu_num == 14:
		occupation = '医療系専門職'
	elif menu_num == 15:
		occupation = 'クリエイター・クリエイティブ職'

	url = "https://doda.jp/"
	# Chrome Driverの設定
	chrome_options = Options()
	# chrome_options.add_argument("--headless")  # ヘッドレスモードで実行
	# Webドライバーの起動
	driver = uc.Chrome(chrome_options)
	driver.maximize_window()
	# 指定サイトにアクセス
	driver.get(url)
	time.sleep(5)

	# サイトの構成が２パターンあるためtry:exceptで使い分け
	try:
		# 職種を選択
		driver.find_element(By.XPATH, '//*[@id="form4"]/div[2]/div/div[1]').click()
		scrollable_element = WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.XPATH, '//*[@id="form4"]/div[2]/div/div[2]/div/div[1]'))
		)
		# クリックしたいターゲット要素を探す
		target_element = WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.XPATH, f'//div[text()="{occupation}"]'))
		)
		while True:
			try:
				# アクションチェーンを使ってターゲット要素まで移動
				actions = ActionChains(driver)
				actions.move_to_element(target_element).perform()
				# クリックする
				target_element.click()
				time.sleep(2)  # 確認用の待機
				break
			except:
				# JavaScriptでスクロールエレメントを操作し、スクロール
				driver.execute_script("arguments[0].scrollTop += 10;", scrollable_element)

		# 勤務地を選択
		driver.find_element(By.XPATH, '//*[@id="form4"]/div[3]/div/div[1]').click()
		# クリックしたいターゲット要素を探す
		target_element = WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.XPATH, f'//div[text()="{work_place}"]'))
		)
		while True:
			try:
				# アクションチェーンを使ってターゲット要素まで移動
				actions = ActionChains(driver)
				actions.move_to_element(target_element).perform()
				# クリックする
				target_element.click()
				time.sleep(2)  # 確認用の待機
				break
			except:
				# JavaScriptでスクロールエレメントを操作し、スクロール
				driver.execute_script("arguments[0].scrollTop += 10;", scrollable_element)

		# 年収を選択
		driver.find_element(By.XPATH, '//*[@id="form4"]/div[4]/div/div[1]').click()
		# クリックしたいターゲット要素を探す
		target_element = WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.XPATH, f'//div[text()="{income}"]'))
		)
		while True:
			try:
				# アクションチェーンを使ってターゲット要素まで移動
				actions = ActionChains(driver)
				actions.move_to_element(target_element).perform()
				# クリックする
				target_element.click()
				time.sleep(2)  # 確認用の待機
				break
			except:
				# JavaScriptでスクロールエレメントを操作し、スクロール
				driver.execute_script("arguments[0].scrollTop += 10;", scrollable_element)

		# フリーワード検索
		driver.find_element(By.XPATH, '//*[@id="k"]').send_keys(free_word)
		time.sleep(2)
		# 検索
		driver.find_element(By.XPATH, '//*[@id="quick_search"]').click()
		time.sleep(5)

	except:
		# 職種を選択
		driver.find_element(By.XPATH, '//*[@id="__next"]/div[2]/main/div[4]/div[1]/div[2]/div/div[1]/button').click()
		scrollable_element = WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/div[2]/main/div[4]/div[1]/div[2]/div/div[1]/div'))
		)
		# クリックしたいターゲット要素を探す
		target_element = WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.XPATH, f'//button[text()="{occupation}"]'))
		)
		while True:
			try:
				# アクションチェーンを使ってターゲット要素まで移動
				actions = ActionChains(driver)
				actions.move_to_element(target_element).perform()
				# クリックする
				target_element.click()
				time.sleep(2)  # 確認用の待機
				break
			except:
				# JavaScriptでスクロールエレメントを操作し、スクロール
				driver.execute_script("arguments[0].scrollTop += 10;", scrollable_element)

		# 勤務地を選択
		driver.find_element(By.XPATH, '//*[@id="__next"]/div[2]/main/div[4]/div[1]/div[2]/div/div[2]/button').click()
		# クリックしたいターゲット要素を探す
		target_element = WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.XPATH, f'//button[text()="{work_place}"]'))
		)
		while True:
			try:
				# アクションチェーンを使ってターゲット要素まで移動
				actions = ActionChains(driver)
				actions.move_to_element(target_element).perform()
				# クリックする
				target_element.click()
				time.sleep(2)  # 確認用の待機
				break
			except:
				# JavaScriptでスクロールエレメントを操作し、スクロール
				driver.execute_script("arguments[0].scrollTop += 10;", scrollable_element)

		# 年収を選択
		driver.find_element(By.XPATH, '//*[@id="__next"]/div[2]/main/div[4]/div[1]/div[2]/div/div[3]/button').click()
		# クリックしたいターゲット要素を探す
		target_element = WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.XPATH, f'//button[text()="{income}"]'))
		)
		while True:
			try:
				# アクションチェーンを使ってターゲット要素まで移動
				actions = ActionChains(driver)
				actions.move_to_element(target_element).perform()
				# クリックする
				target_element.click()
				time.sleep(2)  # 確認用の待機
				break
			except:
				# JavaScriptでスクロールエレメントを操作し、スクロール
				driver.execute_script("arguments[0].scrollTop += 10;", scrollable_element)

		# フリーワード検索
		driver.find_element(By.XPATH, '//*[@id="__next"]/div[2]/main/div[4]/div[1]/div[2]/div/div[4]/div/input').send_keys(free_word)
		time.sleep(2)
		# 検索
		button_element = driver.find_element(By.XPATH, '//*[@id="__next"]/div[2]/main/div[4]/div[1]/div[2]/button')
		driver.execute_script("arguments[0].click();", button_element)
		# button_element.click()
		time.sleep(5)

	# 企業のリスト要素を取得
	company_elements = driver.find_elements(By.XPATH, '//*[@id="__next"]/div[2]/main/div[2]/div[2]/div[2]/div/div/div/div/article/header/div/div[1]/a')  # 企業一覧

	# 各企業のURLをリストに保存
	a_tags = [element.get_attribute('href') for element in company_elements]

	company_list = []
	for a_tag in a_tags:
		driver.get(a_tag)  # URL文字列を取得してアクセス
		time.sleep(2)
		driver.find_element(By.XPATH, '//*[@id="__next"]/main/div[4]/div/div[1]/div[1]/a').click()
		time.sleep(2)
		try:
			company_name = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[2]/div/div/h1').text
			company_name = f'{company_name}\n'
		except NoSuchElementException:
			try:
				company_name = driver.find_element(By.XPATH, '//*[@id="__next"]/main/div[2]/div/hgroup/h1').text
				company_name = f'{company_name}\n'
			except NoSuchElementException:
				company_name = '要素なし(会社名)\n'
		try:
			work = driver.find_element(By.XPATH, '//*[@id="job_description_table"]/tbody/tr[th[text()="仕事内容"]]/td').text
			work = f'{work}\n'
		except NoSuchElementException:
			try:
				work = driver.find_element(By.XPATH, '//*[@id="__next"]/main/div[4]/div/div[2]/div/div[h3[text()="仕事内容"]]/div').text
				work = f'{work}\n'
			except NoSuchElementException:
				work = '要素なし(仕事内容)\n'
		try:
			target_audience = driver.find_element(By.XPATH, '//*[@id="job_description_table"]/tbody/tr[th[text()="対象となる方"]]/td').text
			target_audience = f'{target_audience}\n'
		except NoSuchElementException:
			try:
				target_audience = driver.find_element(By.XPATH, '//*[@id="__next"]/main/div[4]/div/div[2]/div[1]/div/div[h3[text()="対象となる方"]]/div').text
				target_audience = f'{target_audience}\n'
			except NoSuchElementException:
				target_audience = '要素なし(対象者)\n'
		try:
			selection_points = driver.find_element(By.XPATH, '//*[@id="job_description_table"]/tbody/tr[th[text()="選考のポイント"]]/td').text
			selection_points = f'{selection_points}\n'
		except NoSuchElementException:
			try:
				selection_points = driver.find_element(By.XPATH, '//*[@id="__next"]/main/div[4]/div/div[2]/div[1]/div/div[h3[text()="選考のポイント"]]/div').text
				selection_points = f'{selection_points}\n'
			except NoSuchElementException:
				selection_points = '要素なし(選考のポイント)\n'
		try:
			work_place = driver.find_element(By.XPATH, '//*[@id="job_description_table"]/tbody/tr[th[text()="勤務地"]]/td').text
			work_place = f'{work_place}\n'
		except NoSuchElementException:
			try:
				work_place = driver.find_element(By.XPATH, '//*[@id="__next"]/main/div[4]/div/div[2]/div[1]/div/div[h3[text()="勤務地"]]/div').text
				work_place = f'{work_place}\n'
			except NoSuchElementException:
				work_place = '要素なし(勤務地)\n'
		try:
			working_style = driver.find_element(By.XPATH, '//*[@id="job_description_table"]/tbody/tr[th[text()="勤務形態"]]/td').text
			working_style = f'{working_style}\n'
		except NoSuchElementException:
			try:
				working_style = driver.find_element(By.XPATH, '//*[@id="__next"]/main/div[4]/div/div[2]/div[1]/div/div[h3[contains(text(),"勤務形態") or contains(text(),"雇用形態")]]/div').text
				working_style = f'{working_style}\n'
			except NoSuchElementException:
				working_style = '要素なし(勤務形態)\n'
		try:
			salary = driver.find_element(By.XPATH, '//*[@id="job_description_table"]/tbody/tr[th[text()="給与"]]/td').text
			salary = f'{salary}\n'
		except NoSuchElementException:
			try:
				salary = driver.find_element(By.XPATH, '//*[@id="__next"]/main/div[4]/div/div[2]/div[1]/div/div[h3[text()="給与"]]/div').text
				salary = f'{salary}\n'
			except NoSuchElementException:
				salary = '要素なし(給与)\n'
		try:
			welfare_benefits = driver.find_element(By.XPATH, '//*[@id="job_description_table"]/tbody/tr[th[text()="待遇・福利厚生"]]/td').text
			welfare_benefits = f'{welfare_benefits}\n'
		except NoSuchElementException:
			try:
				welfare_benefits = driver.find_element(By.XPATH, '//*[@id="__next"]/main/div[4]/div/div[2]/div[1]/div/div[h3[text()="待遇・福利厚生"]]/div').text
				welfare_benefits = f'{welfare_benefits}\n'
			except NoSuchElementException:
				welfare_benefits = '要素なし(待遇・福利厚生)\n'
		try:
			holiday = driver.find_element(By.XPATH, '//*[@id="job_description_table"]/tbody/tr[th[text()="休日・休暇"]]/td').text
			holiday = f'{holiday}\n'
		except NoSuchElementException:
			try:
				holiday = driver.find_element(By.XPATH, '//*[@id="__next"]/main/div[4]/div/div[2]/div[1]/div/div[h3[text()="休日・休暇"]]/div').text
				holiday = f'{holiday}\n'
			except NoSuchElementException:
				holiday = '要素なし(休日・休暇)\n'

		data_dict = {
			'会社名': company_name,
			'仕事内容': work,
			'対象者': target_audience,
			'選考ポイント': selection_points,
			'勤務地': work_place,
			'勤務形態': working_style,
			'給与': salary,
			'待遇・福利厚生': welfare_benefits,
			'休日休暇': holiday
		}
		company_list.append(data_dict)
		print(f' 会社名: {company_name} データ取得')
	# ブラウザを閉じる
	driver.quit()
	print(' データ取得完了')

	df = pd.DataFrame(company_list)
	spread_sheet(df)
	print(f' スプレッドシート作成完了')


def spread_sheet(df):
    # 2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
	scope = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']
    # 認証情報設定
    # ダウンロードしたjsonファイル名をクレデンシャル変数に設定。Pythonファイルと同じフォルダに置く。
	credentials = Credentials.from_service_account_file(os.getenv('JSON_FILE'), scopes=scope)
    # 共有設定したスプレッドシートキーを格納
	SPREADSHEET_KEY = os.getenv('SPREADSHEET_KEY')
	gc = gspread.authorize(credentials)
	sheet_name = '求人情報'
	sheet = gc.open_by_key(SPREADSHEET_KEY).worksheet(sheet_name)
	# データフレームの内容をリストに変換して書き込み
	sheet.clear()  # シートを一旦クリア（必要に応じて）
	sheet.update([df.columns.values.tolist()] + df.values.tolist())


def main():
	print(' 求人情報を取得します')
	print(' \n 0.プログラムを終了する')
	print(' 1.営業職')
	print(' 2.企画・管理')
	print(' 3.事務・アシスタント')
	print(' 4.販売・サービス職')
	print(' 5.専門職（コンサルティングファーム・専門事務所・監査法人）')
	print(' 6.金融系専門職')
	print(' 7.公務員・教員・農林水産関連職')
	print(' 8.技術職（SE・インフラエンジニア・Webエンジニア）')
	print(' 9.技術職（機械・電気）')
	print(' 10.技術職（組み込みソフトウェア）')
	print(' 11.技術職・専門職（建設・建築・不動産・プラント・工場）')
	print(' 12.技術職（化学・素材・化粧品・トイレタリー）')
	print(' 13.技術職（食品・香料・飼料）')
	print(' 14.医療系専門職')
	print(' 15.クリエイター・クリエイティブ職')
	menu_num = int(input('\n 対象の数字を入力してください(半角): '))
	work_place = str(input('\n 都道府県名を入力してください: '))

	print('\n 年収は200万円〜1000万円まで選択できます。\n 700万円までは50万円単位でそれ以降は100万円単位で条件決められます')
	income = str(input(' 希望年収を入力してください(半角、例:500): '))
	income = f'{income}万円～'
	free_word = str(input('\n スキルや条件を入力してください。もし何もない場合はEnterをクリックしてください。'))

	if menu_num != 0:
		search_Recruitment(menu_num, work_place, income, free_word)
	else:
		print(' プログラムを終了します。')
		sys.exit()


if __name__ == '__main__':
     main()
