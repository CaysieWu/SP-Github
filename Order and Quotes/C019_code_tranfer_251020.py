import pandas as pd
import os
import sqlite3
import cx_Oracle
import re
import numpy as np
from sqlalchemy import create_engine
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

class CODE_TRANSFER:
	def __init__(self, path):
		self.PATH = path
		self.COMPARE_FILE = self.load_compare_file()
		self.Original = None

	# 讀取新舊產品代號比較表   
	def load_compare_file(self):
		compare_path = r"Z:\業務部\業務一課\G-報價\1. 外銷\C01900 Reyher\2025\C019 新舊產品代號對照表_Caysie.xlsx"
		sheet_names = ["MS", "SDS", "TT", "DW", "CB", "TP"]
		dic_item_series = pd.read_excel(compare_path, sheet_name=None, dtype=str)
		return pd.concat([dic_item_series[s] for s in sheet_names], ignore_index=True)

	# 讀取客人原始詢價單或訂單
	def load_customer_file(self):
		for file in os.listdir(self.PATH):
			if "_RFQ_" in file:
				path = os.path.join(self.PATH, file)
				self.Original = pd.read_excel(path, sheet_name="RFQ positions", skiprows=9)
				print(f"Loaded RFQ file: {file}")
				return
			elif "PO_" in file:
				path = os.path.join(self.PATH, file)
				self.Original = pd.read_excel(path, sheet_name="Bestellpositionen", skiprows=1)
				print(f"Loaded PO file: {file}")
				return
		raise FileNotFoundError("No '_RFQ_' or 'PO_' file found in the given path.")

	# 產出舊產品代號
	def find_old_code(self):
		self.Original["old_category"] = np.nan
		self.Original["original_category"] = self.Original["Material"].astype(str).apply(lambda x: ".".join(x.split(".")[:2]))
		for num, item_cat in self.Original["original_category"].items():
			match = self.COMPARE_FILE[self.COMPARE_FILE["新產品代號"] == item_cat]
			if not match.empty:
				self.Original.at[num, "old_category"] = match.iloc[0]["舊產品代號"]
		self.Original["material_suffix"] = self.Original["Material"].astype(str).apply(lambda x: ".".join(x.split(".")[2:]))
		self.Original["previous_item_code"] = self.Original["old_category"].astype(str) + "." + self.Original["material_suffix"]
		return self.Original


class RFQ_CODE_TRANSFER(CODE_TRANSFER):

	def __init__(self, path):
		super().__init__(path)
		self.load_customer_file()
		self.db_data = self.load_db_data()	

	# 讀取所有報價資料庫內的客戶代號為C019的資料
	def load_db_data(self):
		db_path = r"Z:\跨部門\共用資料夾\C. 業務部\詢價統計DB\QUOTATION_DATABASE.db"
		with sqlite3.connect(db_path) as conn:
			return pd.read_sql_query("""
				SELECT PRODUCT_CODE, BOX_TYPE, M_BOX, RECESS_PARTIAL_THREAD
				FROM CUSTOMER_PRODUCT_SUMMARY
				WHERE CUSTOMER_CODE = 'C01900'
				ORDER BY QUOTE_DATE DESC
			""", conn)

	#取得用於跑小工具轉出成本表之最新代號，以新代號優先、舊代號僅次、皆未找到則顯示"Not found"
	def get_search_code(self):
		df = self.find_old_code().copy()
		product_codes = set(self.db_data["PRODUCT_CODE"])
		
		df["複製此欄至小工具"] = "Not found"
		df.loc[df["previous_item_code"].isin(product_codes), "複製此欄至小工具"] = df["previous_item_code"]
		df.loc[df["Material"].isin(product_codes), "複製此欄至小工具"] = df["Material"]
		
		cols = ["Item", "Material", "previous_item_code", "Description", "RFQ Quantity", "Package Qty", "複製此欄至小工具"]
		self.df = df[cols].copy()
		self.output_file = os.path.join(self.PATH, "Updated_RFQ.xlsx")

	#建立找尋ERP資料過去包裝盒型及支數之功能
	def FIND_ERP(self):
		Box_ref = pd.read_excel(r"Z:\業務部\業務一課\Q.工具程式\CAYSIE\Box_ref.xlsx")
		Mapping_box = dict(zip(Box_ref["ERP_CODE"], Box_ref["Quote_Code"]))

		oracle_connection_string = (
			"oracle+cx_oracle://spselect:select@192.168.1.242:1526/?service_name=sperpdb"
		)
		engine = create_engine(oracle_connection_string)

		
		query = f"""
		SELECT "CST_PART_NO", "PDC_1", "DLV_DATE", "PMT_NO", "QTY_PER_CTN"
		FROM ssl_cst_orde_d
		ORDER BY DLV_DATE DESC
		"""
	
		history_data = pd.read_sql_query(query, engine)
		history_data.columns = [col.upper() for col in history_data.columns]
	
		if history_data.empty:
			return None
		history_data = history_data.drop_duplicates(subset=["CST_PART_NO"], keep="first")
		history_data["QTY_PER_CTN"] = history_data["QTY_PER_CTN"] * 1000
		history_data["PMT_NO"] = history_data["PMT_NO"].map(Mapping_box).fillna(history_data["PMT_NO"])
		
		return history_data
	
	#將過去盒型及支數結果新增至Updated_RFQ
	def add_ERP_boxtype(self):

	#以ERP結果優先
		result = self.df.copy()
		
		ERP_feedback = self.FIND_ERP()
		ERP_feedback = ERP_feedback[["CST_PART_NO", "PMT_NO", "QTY_PER_CTN"]]
		ERP_feedback = ERP_feedback.rename(columns={"CST_PART_NO": "latest_erp_code"})
		
		result["latest_erp_code"] = None
		new_code = result["Material"].isin(ERP_feedback["latest_erp_code"])
		result.loc[new_code, "latest_erp_code"] = result.loc[new_code, "Material"]
		old_code = result["previous_item_code"].isin(ERP_feedback["latest_erp_code"])
		missing_code = result["latest_erp_code"].isna() & old_code
		result.loc[missing_code, "latest_erp_code"] = result.loc[missing_code, "previous_item_code"]
		
		result = result.merge(ERP_feedback, how="left", on="latest_erp_code")
		result = result.rename(columns={"PMT_NO": "Previouse_Box"})
		result = result.rename(columns={"QTY_PER_CTN": "Previouse_Box_Quantity"})
		result = result.rename(columns={"latest_erp_code": "資料來源"})
		result.loc[result["資料來源"].notna(), "資料來源"] = "ERP"

		
	#ERP未找到之品項進資料庫尋找
		
		db_data = self.load_db_data()
		db_data = db_data[~db_data["RECESS_PARTIAL_THREAD"].astype(str).str.startswith("NQ")]
		db_data = db_data[["PRODUCT_CODE", "BOX_TYPE", "M_BOX"]].drop_duplicates()
		db_data = db_data.groupby("PRODUCT_CODE").first().reset_index()

		db_merge = result.merge(
			db_data,
			how="left",
			left_on="複製此欄至小工具",
			right_on="PRODUCT_CODE"
		)
		
		missing_mask = result["Previouse_Box"].isna() | (result["Previouse_Box"] == "0001")
		merge_mask = db_merge["BOX_TYPE"].notna() | (db_merge["Previouse_Box"] == "0001")
		update_mask = missing_mask & merge_mask
		
		result.loc[update_mask, "Previouse_Box"] = db_merge.loc[update_mask, "BOX_TYPE"]
		result.loc[update_mask, "Previouse_Box_Quantity"] = db_merge.loc[update_mask, "M_BOX"].astype(float) * 1000
		result.loc[result["Previouse_Box"].notna() & result["資料來源"].isna(), "資料來源"] = "資料庫"
		result.loc[result["Previouse_Box"].isna(), "Previouse_Box_Quantity"] = ""
		result.loc[result["Previouse_Box"].isna(), "資料來源"] = ""

	#新增欄位，比對結果支數及詢價支數以確保資料正確性，不同者會以紅色底色顯示
		with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
			result.to_excel(writer, index=False, sheet_name='Sheet1')
			workbook = writer.book
			worksheet = writer.sheets['Sheet1']

			red_format = workbook.add_format({'bg_color': '#FF0000'})

			for row in range(1, len(result) + 1):  # +1 to offset header
				pkg_cell = f'F{row + 1}'  # +1 because Excel rows are 1-based instead of 0-based
				prev_pkg__cell = f'J{row + 1}'
				worksheet.conditional_format(prev_pkg__cell, {
					'type': 'formula',
					'criteria': f'={pkg_cell}<>{prev_pkg__cell}',
					'format': red_format
				})


class PO_CODE_TRANSFER(CODE_TRANSFER):
	def __init__(self, path):
		super().__init__(path)
		self.load_customer_file() #確保啟動功能時，客戶原始檔案同步讀取
		self.customer_items = self.fetch_customer_items()

	#連結ERP 0605M 取得曾出貨且客戶代號為C01900之所有客戶產品代號 (0210M尚未有table可拉)
	def fetch_customer_items(self):
		oracle_connection_string = (
			"oracle+cx_oracle://spselect:select@192.168.1.242:1526/?service_name=sperpdb"
		)
		engine = create_engine(oracle_connection_string)
		df = pd.read_sql_query("SELECT CST_PART_NO FROM v_ssl0605q_cst WHERE ORD_CST_NO = 'C01900'", engine)
		
		df.columns = [col.upper() for col in df.columns]
		return df.drop_duplicates(subset=["CST_PART_NO"])
	
	#進資料庫尋找各品項的最新報價
	def load_db_data(self):
		db_path = r"Z:\跨部門\共用資料夾\C. 業務部\詢價統計DB\QUOTATION_DATABASE.db"
		with sqlite3.connect(db_path) as conn:
			return pd.read_sql_query("""
				WITH LatestQuote AS (
				SELECT *,
					   ROW_NUMBER() OVER (
						   PARTITION BY PRODUCT_CODE
						   ORDER BY QUOTE_DATE DESC
					   ) as rn
				FROM CUSTOMER_PRODUCT_SUMMARY
				WHERE CUSTOMER_CODE = 'C01900'
			)
			SELECT PRODUCT_CODE, INTEREST_GROSS, COLLATED_LOSS, COLLATED_PRICE, EXCHANGE_RATE, TOTAL_PRICE_M
			FROM LatestQuote
			WHERE rn = 1
			""", conn)
	
	#根據資料庫最新資料，計算實際成本。
	def count_cost(self):
		db_data = self.load_db_data().copy()
		order_item = self.Original["Material"]

		db_data['PRODUCT_CODE'] = db_data['PRODUCT_CODE'].str.strip()
		for col in db_data.columns:
			if col != 'PRODUCT_CODE':
				db_data[col] = pd.to_numeric(db_data[col], errors='coerce').fillna(0).round(2)
		
		costs = []
		for item in order_item:
			matched_row = db_data[db_data['PRODUCT_CODE'] == item]
			if not matched_row.empty:
				row = matched_row.iloc[0]
				rebate_cost = row["TOTAL_PRICE_M"] * 0.025 * row["EXCHANGE_RATE"]
				collated_cost = row["COLLATED_PRICE"] * (1 + row["COLLATED_LOSS"])
				total_cost = row["INTEREST_GROSS"] + rebate_cost + collated_cost

				costs.append({"PRODUCT_CODE": item, "TOTAL_COST": round(total_cost, 2)})
			else:
				costs.append({"PRODUCT_CODE": item, "TOTAL_COST": None})
		self.cost_df = pd.DataFrame(costs)
				

	# 比對0605M後，若客人已以新代號下過單顯示"New Item Code exist"；
	# 若無則顯示舊產品代號；舊產品代號不存在時顯示"初次下單"
	# 加入成本單價NTD/M
	def set_item_code(self):
		df = self.find_old_code().copy()
		new_code = df["Material"].isin(self.customer_items["CST_PART_NO"])
		old_code = df["previous_item_code"].isin(self.customer_items["CST_PART_NO"])
		df["set_code"] = "初次下單"
		df.loc[old_code, "set_code"] = df.loc[old_code, "previous_item_code"]
		df.loc[new_code, "set_code"] = "New Item Code exist"

		df["成本單價NTD/M"] = df["Material"].map(self.cost_df.set_index("PRODUCT_CODE")["TOTAL_COST"])

		cols = ["Purchasing Doc.", "Item", "Material", "set_code", "Description 2", "PO Quantity", "成本單價NTD/M"]
		df[cols].to_excel(os.path.join(self.PATH, "Set_Item_Code.xlsx"), index=False)



if __name__ == "__main__":
	path = input("Enter folder path: ").strip()

	#根據路徑判斷要啟動之功能
	if "報價" in path:
		rfq_bot = RFQ_CODE_TRANSFER(path)
		rfq_bot.get_search_code()
		rfq_bot.add_ERP_boxtype()
		print("Updated_RFQ saved in path.")

	elif "訂單" in path:
		po_bot = PO_CODE_TRANSFER(path)
		po_bot.load_db_data()
		po_bot.count_cost()
		po_bot.set_item_code()
		print("Check to-be-set-codes file in path.")

	else:
		print("Unable to define '報價' or '訂單'")

	
