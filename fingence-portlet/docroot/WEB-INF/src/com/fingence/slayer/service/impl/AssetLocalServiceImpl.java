/**
 * Copyright (c) 2000-2013 Liferay, Inc. All rights reserved.
 *
 * This library is free software; you can redistribute it and/or modify it under
 * the terms of the GNU Lesser General Public License as published by the Free
 * Software Foundation; either version 2.1 of the License, or (at your option)
 * any later version.
 *
 * This library is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU Lesser General Public License for more
 * details.
 */

package com.fingence.slayer.service.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fingence.IConstants;
import com.fingence.slayer.NoSuchAssetException;
import com.fingence.slayer.NoSuchBondException;
import com.fingence.slayer.NoSuchEquityException;
import com.fingence.slayer.NoSuchMutualFundException;
import com.fingence.slayer.model.Asset;
import com.fingence.slayer.model.Bond;
import com.fingence.slayer.model.Dividend;
import com.fingence.slayer.model.Equity;
import com.fingence.slayer.model.History;
import com.fingence.slayer.model.MutualFund;
import com.fingence.slayer.service.base.AssetLocalServiceBaseImpl;
import com.fingence.util.CellUtil;
import com.liferay.portal.kernel.exception.SystemException;
import com.liferay.portal.kernel.log.Log;
import com.liferay.portal.kernel.log.LogFactoryUtil;
import com.liferay.portal.kernel.util.Validator;
import com.liferay.portal.model.Country;
import com.liferay.portal.security.auth.CompanyThreadLocal;
import com.liferay.portal.service.CountryServiceUtil;
import com.liferay.portal.service.ServiceContext;

/**
 * The implementation of the asset local service.
 * 
 * <p>
 * All custom service methods should be put in this class. Whenever methods are
 * added, rerun ServiceBuilder to copy their definitions into the
 * {@link com.fingence.slayer.service.AssetLocalService} interface.
 * 
 * <p>
 * This is a local service. Methods of this service will not have security
 * checks based on the propagated JAAS credentials because this service can only
 * be accessed from within the same VM.
 * </p>
 * 
 * @author Ahmed Hasan
 * @see com.fingence.slayer.service.base.AssetLocalServiceBaseImpl
 * @see com.fingence.slayer.service.AssetLocalServiceUtil
 */
public class AssetLocalServiceImpl extends AssetLocalServiceBaseImpl {
	/*
	 * NOTE FOR DEVELOPERS:
	 * 
	 * Never reference this interface directly. Always use {@link
	 * com.fingence.slayer.service.AssetLocalServiceUtil} to access the asset
	 * local service.
	 */

	public void loadPricingData(long userId, File excelFile, ServiceContext serviceContext, int type) {
		
		System.out.println("inside Load Pricing Data....");
		if (Validator.isNull(excelFile)) return;
		
		InputStream is = null;
		try {
			is = new FileInputStream(excelFile);
		} catch (FileNotFoundException e) {
			//e.printStackTrace();
		}
				
		if (Validator.isNull(is)) return;
		
		// Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(is);
		} catch (IOException e) {
			e.printStackTrace();
		}

		if (Validator.isNull(workbook)) return;

		// Get first/desired sheet from the workbook
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		Iterator<Row> rowIterator = sheet.iterator();
		Map<Integer, Long> columnNames = new HashMap<Integer, Long>();
		int columnCount = 0;
		
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			
			columnCount = row.getPhysicalNumberOfCells();
			
			_log.debug("processing row ==> " + row.getRowNum());
			System.out.println("processing row ==> " + row.getRowNum());
			
			int rowNum = row.getRowNum();
			
			if (rowNum == 0) continue;
			
			if (rowNum == 1) {
				for (int i=0; i < columnCount; i++){
					Cell cell = row.getCell(i);
					if (Validator.isNull(cell)) continue;
					
					String id_isin = CellUtil.getString(cell);
										
					Asset asset = null;
					try {
						asset = assetPersistence.fetchByIdISIN(id_isin);
					} catch (SystemException e) {
						e.printStackTrace();
					}
					
					if (Validator.isNull(asset)) continue;
					
					columnNames.put(i, asset.getAssetId());
				}
				continue;
			}
			
			if (rowNum > 1 && rowNum < 14) continue; 
						
			System.out.println("going to process data...");
			
			Iterator<Integer> itr = columnNames.keySet().iterator();
						
			//for (int i=3; i < columnCount; i++){
			
			while (itr.hasNext()) {
				
				int i = itr.next();
				Date date = CellUtil.getDate(row.getCell(i));
				
				if (Validator.isNull(date)) continue;
					
				long assetId = 0l;
				try {
					assetId = columnNames.get(i);
				} catch (Exception e) {
					_log.debug(e.getMessage() + ": There is an exception...");
					continue;
				}
				
				double value = CellUtil.getDouble(row.getCell(++i));
				
				History history = null;
				try {
					history = historyPersistence.fetchByAssetId_Date_Type(assetId, date, type);
					_log.debug("history record already present...");
				} catch (SystemException e) {
					e.printStackTrace();
				}
				
				if (Validator.isNull(history)) {
					long recId = 0l;
					try {
						recId = counterLocalService.increment(History.class.getName());
					} catch (SystemException e) {
						e.printStackTrace();
					}
					
					history = historyLocalService.createHistory(recId);
					try {
						history = historyLocalService.addHistory(history);
					} catch (SystemException e) {
						e.printStackTrace();
					}
				}
				
				history.setAssetId(assetId);
				history.setType(type);
				history.setValue(value);
				history.setLogDate(date);
				
				if (type == IConstants.HISTORY_TYPE_BOND_CASHFLOW) {
					double principal = CellUtil.getDouble(row.getCell(++i));
					history.setPrincipal(principal);
				}
				
				try {
					history = historyLocalService.updateHistory(history);
				} catch (SystemException e) {
					e.printStackTrace();
				}	
			}
		}
	}
	
	public void loadDividends(long userId, File excelFile, ServiceContext serviceContext) {
		
		System.out.println("inside Load Dividends Data....");
		if (Validator.isNull(excelFile)) return;
		
		InputStream is = null;
		try {
			is = new FileInputStream(excelFile);
		} catch (FileNotFoundException e) {
			//e.printStackTrace();
		}
				
		if (Validator.isNull(is)) return;
		
		// Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(is);
		} catch (IOException e) {
			e.printStackTrace();
		}

		if (Validator.isNull(workbook)) return;

		// Get first/desired sheet from the workbook
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		Iterator<Row> rowIterator = sheet.iterator();
		Map<Integer, Long> columnNames = new HashMap<Integer, Long>();
		int columnCount = 0;
		
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			
			columnCount = row.getPhysicalNumberOfCells();
			
			_log.debug("processing row ==> " + row.getRowNum());
			System.out.println("processing row ==> " + row.getRowNum());
			
			if (row.getRowNum() == 0) continue;
			
			if (row.getRowNum() == 1) {
				for (int i=0; i < columnCount; i++){
					Cell cell = row.getCell(i);
					if (Validator.isNull(cell)) continue;
					
					String id_isin = CellUtil.getString(cell);
										
					Asset asset = null;
					try {
						asset = assetPersistence.fetchByIdISIN(id_isin);
					} catch (SystemException e) {
						e.printStackTrace();
					}
					
					if (Validator.isNull(asset)) continue;
					
					columnNames.put(i, asset.getAssetId());
				}
				continue;
			}
						
			for (int i=0; i < columnCount; i++){
				Date declaredDate = CellUtil.getDate(row.getCell(i));
				
				if (Validator.isNull(declaredDate)) continue;
					
				long assetId = 0l;
				try {
					assetId = columnNames.get(i);
				} catch (Exception e) {
					_log.debug(e.getMessage() + ": There is an exception...");
					continue;
				}
				
				Date exDate = CellUtil.getDate(row.getCell(++i));
				
				Date recordDate = CellUtil.getDate(row.getCell(++i));
				
				Date payableDate = CellUtil.getDate(row.getCell(++i));
				
				double amount = CellUtil.getDouble(row.getCell(++i));
				String frequency = CellUtil.getString(row.getCell(++i));
				String type = CellUtil.getString(row.getCell(++i));
				
				Dividend dividend = null;
				try {
					dividend = dividendPersistence.fetchByAssetId_DeclaredDate(assetId, declaredDate);
					_log.debug("dividend record already present...");
				} catch (SystemException e) {
					e.printStackTrace();
				}
				
				if (Validator.isNull(dividend)) {
					
					long recId = 0l;
					try {
						recId = counterLocalService.increment(Dividend.class.getName());
					} catch (SystemException e) {
						e.printStackTrace();
					}
					
					dividend = dividendPersistence.create(recId);
				}	
				
				// update the record
				dividend.setDeclaredDate(declaredDate);
				dividend.setExDate(exDate);
				dividend.setAssetId(assetId);
				dividend.setRecordDate(recordDate);
				dividend.setPayableDate(payableDate);
				dividend.setAmount(amount);
				dividend.setFrequency(frequency);
				dividend.setType(type);
				
				try {
					dividend = dividendLocalService.updateDividend(dividend);
					System.out.println("dividend new history records..." + dividend);
				} catch (SystemException e) {
					e.printStackTrace();
				}
			}
		}
	}
	
	public void importFromExcel(long userId, File excelFile, ServiceContext serviceContext) {

		if (Validator.isNull(excelFile)) return;
		
		InputStream is = null;
		try {
			is = new FileInputStream(excelFile);
		} catch (FileNotFoundException e) {
			//e.printStackTrace();
		}
				
		if (Validator.isNull(is)) return;
		
		// Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(is);
		} catch (IOException e) {
			e.printStackTrace();
		}

		if (Validator.isNull(workbook)) return;

		// Get first/desired sheet from the workbook
		XSSFSheet sheet = workbook.getSheetAt(0);

		// Iterate through each rows one by one
		Iterator<Row> rowIterator = sheet.iterator();
		Map<String, Integer> columnNames = new HashMap<String, Integer>();
		int columnCount = 0;
		
		long bbSecurityVocabularyId = AssetHelper.getVocabularyId(userId, "BB_Security", serviceContext);
		long bbIndustryVocabularyId = AssetHelper.getVocabularyId(userId, "BB_Industry", serviceContext);
		long bbAssetClassVocabularyId = AssetHelper.getVocabularyId(userId, "BB_Asset_Class", serviceContext);
		
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
						
			columnCount = row.getPhysicalNumberOfCells();
			
			if (row.getRowNum() == 0) {
				for (int i=0; i < columnCount; i++){
					Cell cell = row.getCell(i);
					if (Validator.isNotNull(cell)) {
						columnNames.put(CellUtil.getStringCaps(cell), i);
					}
				}
				continue;
			}

			String id_isin = CellUtil.getString(row.getCell(columnNames.get("ID_ISIN")));
			
			if (Validator.isNull(id_isin)) {
				System.out.println("id_isin is null or empty.. continuing...the name is..." + CellUtil.getString(row.getCell(columnNames.get("NAME"))));
				continue;
			}
			
			Asset asset = getAsset(userId, id_isin);
			
			asset.setSecurity_ticker(CellUtil.getString(row.getCell(columnNames.get("SECURITY_TICKER"))));
			asset.setId_cusip(CellUtil.getString(row.getCell(columnNames.get("ID_CUSIP"))));
			asset.setId_bb_global(CellUtil.getString(row.getCell(columnNames.get("ID_BB_GLOBAL"))));
			asset.setId_bb_sec_num_src(CellUtil.getLong(row.getCell(columnNames.get("ID_BB_SEC_NUM_SRC"))));
			asset.setName(CellUtil.getString(row.getCell(columnNames.get("NAME"))));
			asset.setChg_pct_mtd(CellUtil.getDouble(row.getCell(columnNames.get("CHG_PCT_MTD"))));
			asset.setChg_pct_5d(CellUtil.getDouble(row.getCell(columnNames.get("CHG_PCT_5D"))));
			asset.setChg_pct_1m(CellUtil.getDouble(row.getCell(columnNames.get("CHG_PCT_1M"))));
			asset.setChg_pct_3m(CellUtil.getDouble(row.getCell(columnNames.get("CHG_PCT_3M"))));
			asset.setChg_pct_6m(CellUtil.getDouble(row.getCell(columnNames.get("CHG_PCT_6M"))));
			asset.setChg_pct_ytd(CellUtil.getDouble(row.getCell(columnNames.get("CHG_PCT_YTD"))));
			asset.setBid_price(CellUtil.getDouble(row.getCell(columnNames.get("PX_BID"))));
			asset.setAsk_price(CellUtil.getDouble(row.getCell(columnNames.get("PX_ASK"))));
			asset.setLast_price(CellUtil.getDouble(row.getCell(columnNames.get("PX_LAST"))));
			asset.setChg_pct_high_52week(CellUtil.getDouble(row.getCell(columnNames.get("CHG_PCT_HIGH_52WEEK"))));
			asset.setChg_pct_low_52week(CellUtil.getDouble(row.getCell(columnNames.get("CHG_PCT_LOW_52WEEK"))));
			asset.setSecurity_des(CellUtil.getString(row.getCell(columnNames.get("SECURITY_DES"))));
			asset.setParent_comp_name(CellUtil.getString(row.getCell(columnNames.get("PARENT_COMP_NAME"))));
			
			String securityClass = CellUtil.getString(row.getCell(columnNames.get("BPIPE_REFERENCE_SECURITY_CLASS")));
			
			if (securityClass.equalsIgnoreCase("FixedIncome")) {
				securityClass = "Fixed Income";
			}
			
			asset.setVolatility_30d(CellUtil.getDouble(row.getCell(columnNames.get("VOLATILITY_30D"))));
			asset.setVolatility_90d(CellUtil.getDouble(row.getCell(columnNames.get("VOLATILITY_90D"))));
			asset.setVolatility_180d(CellUtil.getDouble(row.getCell(columnNames.get("VOLATILITY_180D"))));
			asset.setVolatility_360d(CellUtil.getDouble(row.getCell(columnNames.get("VOLATILITY_360D"))));
			
			asset.setCurrency(CellUtil.getString(row.getCell(columnNames.get("CRNCY"))).toUpperCase());
			
			Country country = null;
			try {
				String countryCode = CellUtil.getString(row.getCell(columnNames.get("CNTRY_OF_DOMICILE")));
				
				if (countryCode.equalsIgnoreCase("SP")) {
					countryCode = "ES";
				} else if (countryCode.equalsIgnoreCase("EN")) {
					countryCode = "GB";
				}
				country = CountryServiceUtil.fetchCountryByA2(countryCode);
			} catch (SystemException e) {
				e.printStackTrace();
			}
			if (Validator.isNotNull(country)) {
				asset.setCountry(country.getCountryId());
			}
			
			country = null;
			try {
				//CNTRY_OF_RISK
				String countryCode = CellUtil.getString(row.getCell(columnNames.get("CNTRY_OF_RISK")));
				
				if (countryCode.equalsIgnoreCase("SP")) {
					countryCode = "ES";
				} else if (countryCode.equalsIgnoreCase("EN")) {
					countryCode = "GB";
				}				
				
				country = CountryServiceUtil.fetchCountryByA2(countryCode);
			} catch (SystemException e) {
				e.printStackTrace();
			}
			
			if (Validator.isNotNull(country)) {
				asset.setCountry_of_risk(country.getCountryId());
			} else {
				asset.setCountry_of_risk(asset.getCountry());
			}
			
			if (securityClass.equalsIgnoreCase("Fixed Income")) {
				asset.setSecurity_class(IConstants.SECURITY_CLASS_FIXED_INCOME);
				asset.setCurrent_price(asset.getBid_price()/100);
			} else if (securityClass.equalsIgnoreCase("Fund")) {
				asset.setSecurity_class(IConstants.SECURITY_CLASS_FUND);
				asset.setCurrent_price(CellUtil.getDouble(row.getCell(columnNames.get("FUND_NET_ASSET_VAL"))));
			} else {
				asset.setSecurity_class(IConstants.SECURITY_CLASS_EQUITY);
				asset.setCurrent_price(asset.getLast_price());
			}
			
			try {
				updateAsset(asset);
			} catch (SystemException e) {
				e.printStackTrace();
			}
			
			long assetId = asset.getAssetId();
			
			// Saving to AssetEntry table
			long entryId = AssetHelper.updateAssetEntry(assetId);
			
			AssetHelper.assignCategories(asset, entryId, userId, row,
					columnNames, serviceContext, bbSecurityVocabularyId,
					bbIndustryVocabularyId, bbAssetClassVocabularyId);
			
			if (securityClass.equalsIgnoreCase("Fixed Income")) {
				Bond bond = getBond(assetId);
             	bond.setIssuer_bulk(CellUtil.getString(row.getCell(columnNames.get("ISSUER_BULK"))));
             	bond.setCpn(CellUtil.getDouble(row.getCell(columnNames.get("CPN"))));
             	bond.setCpn_typ(CellUtil.getString(row.getCell(columnNames.get("CPN_TYP"))));
             	bond.setMty_typ(CellUtil.getString(row.getCell(columnNames.get("MTY_TYP"))));
             	bond.setMty_years_tdy(CellUtil.getDouble(row.getCell(columnNames.get("MTY_YEARS_TDY"))));
             	bond.setYld_ytm_ask(CellUtil.getDouble(row.getCell(columnNames.get("YLD_YTM_ASK"))));
             	bond.setYld_ytm_bid(CellUtil.getDouble(row.getCell(columnNames.get("YLD_YTM_BID"))));
             	bond.setYld_cur_mid(CellUtil.getDouble(row.getCell(columnNames.get("YLD_CUR_MID"))));
             	bond.setBb_composite(CellUtil.getString(row.getCell(columnNames.get("BB_COMPOSITE"))));
            	bond.setRtg_sp(CellUtil.getString(row.getCell(columnNames.get("RTG_SP"))));
            	bond.setRtg_moody(CellUtil.getString(row.getCell(columnNames.get("RTG_MOODY"))));
            	bond.setRtg_fitch(CellUtil.getString(row.getCell(columnNames.get("RTG_FITCH"))));
            	bond.setCpn_freq(CellUtil.getDouble(row.getCell(columnNames.get("CPN_FREQ"))));
            	bond.setFive_y_bid_cds_spread(CellUtil.getDouble(row.getCell(columnNames.get("5Y_BID_CDS_SPREAD"))));
            	bond.setDur_mid(CellUtil.getDouble(row.getCell(columnNames.get("DUR_MID"))));
             	bond.setPrice_to_cash_flow(CellUtil.getDouble(row.getCell(columnNames.get("PX_TO_CASH_FLOW"))));
             	bond.setMaturity_dt(CellUtil.getDate(row.getCell(columnNames.get("MATURITY"))));
             	bond.setCollat_typ(CellUtil.getString(row.getCell(columnNames.get("PAYMENT_RANK"))));
             	bond.setCalc_typ(CellUtil.getDouble(row.getCell(columnNames.get("CALC_TYP"))));
             	bond.setIs_bond_no_calctyp(Validator.isNull(CellUtil.getString(row.getCell(columnNames.get("IS_BOND_NO_CALCTYP")))));
             	bond.setIssue_dt(CellUtil.getDate(row.getCell(columnNames.get("ISSUE_DT"))));
             	bond.setAmount_issued(CellUtil.getDouble(row.getCell(columnNames.get("AMT_ISSUED"))));
             	bond.setAmount_outstanding(CellUtil.getDouble(row.getCell(columnNames.get("AMT_OUTSTANDING"))));
             		
             	try {
					bondLocalService.updateBond(bond);
				} catch (SystemException e) {
					e.printStackTrace();
				}
				
			} else if (securityClass.equalsIgnoreCase("Fund")) {
				MutualFund mutualFund = getMutualFund(assetId);
             	mutualFund.setFund_total_assets(CellUtil.getDouble(row.getCell(columnNames.get("FUND_TOTAL_ASSETS"))));
             	mutualFund.setFund_asset_class_focus(CellUtil.getString(row.getCell(columnNames.get("FUND_ASSET_CLASS_FOCUS"))));
             	mutualFund.setFund_geo_focus(CellUtil.getString(row.getCell(columnNames.get("FUND_GEO_FOCUS"))));
             	
             	try {
					mutualFundLocalService.updateMutualFund(mutualFund);
				} catch (SystemException e) {
					e.printStackTrace();
				}
				
			} else if (securityClass.equalsIgnoreCase("Equity")) {
				Equity equity = getEquity(assetId);
             	equity.setEqy_alpha(CellUtil.getDouble(row.getCell(columnNames.get("EQY_ALPHA"))));
             	equity.setDividend_yield(CellUtil.getDouble(row.getCell(columnNames.get("DIVIDEND_YIELD"))));
             	equity.setEqy_dvd_yld_12m(CellUtil.getDouble(row.getCell(columnNames.get("EQY_DVD_YLD_12M"))));
             	equity.setEqy_dvd_yld_es(CellUtil.getDouble(row.getCell(columnNames.get("EQY_DVD_YLD_EST"))));
             	equity.setDvd_payout_ratio(CellUtil.getDouble(row.getCell(columnNames.get("DVD_PAYOUT_RATIO"))));
             	equity.setPe_ratio(CellUtil.getDouble(row.getCell(columnNames.get("PE_RATIO"))));
             	equity.setTot_debt_to_com_eqy(CellUtil.getDouble(row.getCell(columnNames.get("TOT_DEBT_TO_COM_EQY"))));
             	equity.setEbitda_to_revenue(CellUtil.getDouble(row.getCell(columnNames.get("EBITDA_TO_REVENUE"))));
             	equity.setTrail_12m_prof_margin(CellUtil.getDouble(row.getCell(columnNames.get("TRAIL_12M_PROF_MARGIN"))));
             	equity.setBest_current_ev_best_opp(CellUtil.getDouble(row.getCell(columnNames.get("BEST_CURRENT_EV_BEST_OPP"))));
             	equity.setEqy_beta(CellUtil.getDouble(row.getCell(columnNames.get("EQY_ALPHA"))));
              	equity.setReturn_sharpe_ratio(CellUtil.getDouble(row.getCell(columnNames.get("RETURN_SHARPE_RATIO"))));
             	equity.setEqy_sharpe_ratio_1yr(CellUtil.getDouble(row.getCell(columnNames.get("EQY_SHARPE_RATIO_1YR"))));
             	equity.setEqy_sharpe_ratio_3yr(CellUtil.getDouble(row.getCell(columnNames.get("EQY_SHARPE_RATIO_3YR"))));
             	equity.setEqy_sharpe_ratio_5yr(CellUtil.getDouble(row.getCell(columnNames.get("EQY_SHARPE_RATIO_5YR"))));
             	
             	try {
					equityLocalService.updateEquity(equity);
				} catch (SystemException e) {
					e.printStackTrace();
				}
			}
		}		
	}

	private Equity getEquity(long assetId) {
		Equity equity = null;
		try {
			equity = equityPersistence.findByPrimaryKey(assetId);
		} catch (NoSuchEquityException e) {
			equity = equityLocalService.createEquity(assetId);
			try {
				equity = equityLocalService.addEquity(equity);
			} catch (SystemException e1) {
				e1.printStackTrace();
			}
		} catch (SystemException e) {
			e.printStackTrace();
		}
		
		return equity;
	}

	private MutualFund getMutualFund(long assetId) {
		MutualFund mutualFund = null;
		try {
			mutualFund = mutualFundPersistence.findByPrimaryKey(assetId);
		} catch (NoSuchMutualFundException e) {
			mutualFund = mutualFundLocalService.createMutualFund(assetId);
			try {
				mutualFund = mutualFundLocalService.addMutualFund(mutualFund);
			} catch (SystemException e1) {
				e1.printStackTrace();
			}
		} catch (SystemException e) {
			e.printStackTrace();
		}
		
		return mutualFund;
	}

	private Bond getBond(long assetId) {
		
		Bond bond = null;
		try {
			bond = bondPersistence.findByPrimaryKey(assetId);
		} catch (NoSuchBondException e) {
			bond = bondLocalService.createBond(assetId);
			try {
				bond = bondLocalService.addBond(bond);
			} catch (SystemException e1) {
				e1.printStackTrace();
			}
		} catch (SystemException e) {
			e.printStackTrace();
		}
		
		return bond;
	}

	private Asset getAsset(long userId, String id_isin) {
		
		long companyId = CompanyThreadLocal.getCompanyId();
		
		Asset asset = null;
		
		long assetId = getAssetId(id_isin);
		
		if (assetId > 0l) {
			try {
				asset = fetchAsset(assetId);
				asset.setModifiedDate(new java.util.Date());
			} catch (SystemException e) {
				e.printStackTrace();
			}
		} else {
			try {
				assetId = counterLocalService.increment(Asset.class.getName());
				asset = createAsset(assetId);
				asset.setCompanyId(companyId);
				asset.setUserId(userId);
				asset.setCreateDate(new java.util.Date());
				asset = addAsset(asset);
			} catch (SystemException e) {
				e.printStackTrace();
			}
		}
		
		asset.setId_isin(id_isin);
		
		return asset;
	}

	private long getAssetId(String id_isin) {
		
		long assetId = 0l;
		
		try {
			Asset asset = assetPersistence.findByIdISIN(id_isin);
			
			if (Validator.isNotNull(asset)){
				assetId = asset.getAssetId();
			}
		} catch (NoSuchAssetException e) {
			System.out.println("No Asset with this id found..." + id_isin);
			//e.printStackTrace();
		} catch (SystemException e) {
			//e.printStackTrace();
		}
		
		return assetId;
	}
	
	private static final Log _log = LogFactoryUtil.getLog(AssetLocalServiceImpl.class);
}