
################################################################################################################
############## 								ESTRAE TUTTE LE CONFIGURAZIONI DI MARGINE								     ##############
################################################################################################################
SELECT
MARGIN_CONFIGURATION.ID
,ID_BUSINESS_PROFILE
,ID_AIRLINE_SUPPLIER_CODE
,ROUTE
,ID_GLOBAL_COMMISSION              
,FLAT_AMOUNT     
,MIN_AMOUNT       
,MAX_AMOUNT        
,PERCENTAGE_AMOUNT
,MARGIN_COMMISSION_RULE.TYPE            
,MATRIX_GROUP
,ID_MARGIN_PARTITION
,FLIGHT_COMMISSION
,AGENCY_COMMISSION
,SUPPLEMENT_COMMISSION
,ID_DISCOUNT
,FLAT_DISCOUNT
,MIN_DISCOUNT
,MAX_DISCOUNT
,PERCENTAGE_DISCOUNT
FROM MARGIN_CONFIGURATION
LEFT JOIN MARGIN_COMMISSION_RULE ON MARGIN_CONFIGURATION.id_global_commission=MARGIN_COMMISSION_RULE.id
LEFT JOIN MARGIN_DISCOUNT_RULE ON MARGIN_CONFIGURATION.id_discount=MARGIN_DISCOUNT_RULE.id
LEFT JOIN MARGIN_PARTITION ON MARGIN_CONFIGURATION.id_margin_partition=MARGIN_PARTITION.id;

############## ESTRAE TUTTE LE CONFIGURAZIONI DI MARGINE FORMATTATO PER CREARE IL FILE CSV PER IMPORT MASSIVO
SELECT
ID_BUSINESS_PROFILE AS 'BUSINESS_PROFILE'
,ID_AIRLINE_SUPPLIER_CODE AS 'AIRLINE'
,ROUTE           
,REPLACE(FLAT_AMOUNT,'.',',') AS 'FLAT_AMOUNT'     
,REPLACE(MIN_AMOUNT,'.',',') AS 'MIN_AMOUNT'       
,REPLACE(MAX_AMOUNT,'.',',') AS 'MAX_AMOUNT'        
,REPLACE(PERCENTAGE_AMOUNT,'.',',') AS 'PERCENTAGE_AMOUNT'          
,IFNULL(MATRIX_GROUP,'') AS 'MARGIN_CORRECTION_MATRIX_GROUP'
,REPLACE(FLAT_DISCOUNT,'.',',') AS 'FLAT_DISCOUNT'
,REPLACE(MIN_DISCOUNT,'.',',') AS 'MIN_DISCOUNT'
,REPLACE(MAX_DISCOUNT,'.',',') AS 'MAX_DISCOUNT'
,REPLACE(PERCENTAGE_DISCOUNT,'.',',') AS 'PERCENTAGE_DISCOUNT'
,REPLACE(FLIGHT_COMMISSION,'.',',') AS 'FLIGHT_COMMISSION'
,REPLACE(AGENCY_COMMISSION,'.',',') AS 'AGENCY_COMMISSION'
,REPLACE(SUPPLEMENT_COMMISSION,'.',',') AS 'SUPPLEMENT_COMMISSION'
,"0" AS 'DELETE'
FROM MARGIN_CONFIGURATION
LEFT JOIN MARGIN_COMMISSION_RULE ON MARGIN_CONFIGURATION.id_global_commission=MARGIN_COMMISSION_RULE.id
LEFT JOIN MARGIN_DISCOUNT_RULE ON MARGIN_CONFIGURATION.id_discount=MARGIN_DISCOUNT_RULE.id
LEFT JOIN MARGIN_PARTITION ON MARGIN_CONFIGURATION.id_margin_partition=MARGIN_PARTITION.id
WHERE ID_BUSINESS_PROFILE='XXXXXX';



################################################################################################################
############## 								ESTRAE TUTTE LE MATRICI CORRETTIVE										     ##############
################################################################################################################
SELECT
ID_MATRIX_GROUP,
SEGMENT,
RANGE_MIN,
RANGE_MAX,
AMOUNT,
TYPE
FROM MARGIN_CORRECTION_MATRIX; 

############## ESTRAE TUTTE LE MATRICI CORRETTIVE FORMATTATE PER CREARE IL FILE CSV PER IMPORT MASSIVO
SELECT
ID_MATRIX_GROUP,
DESCRIPTION,
SEGMENT AS 'SEGMENTS',
REPLACE(RANGE_MIN,'.',',') AS 'RANGE_MIN',     
REPLACE(RANGE_MAX,'.',',') AS 'RANGE_MAX',       
REPLACE(AMOUNT,'.',',') AS 'AMOUNT',  
TYPE
FROM MARGIN_CORRECTION_MATRIX
LEFT JOIN MARGIN_CORRECTION_MATRIX_GROUP ON MARGIN_CORRECTION_MATRIX.id_matrix_group=MARGIN_CORRECTION_MATRIX_GROUP.id; 



################################################################################################################
############## 								ESTRAE TUTTE LE CONFIGURAZIONI DI PAGAMENTO							     ##############
################################################################################################################
SELECT PAYMENT_METHODS_PACKAGE.*
,MARGIN_COMMISSION_RULE.FLAT_AMOUNT     
,MARGIN_COMMISSION_RULE.MIN_AMOUNT     
,MARGIN_COMMISSION_RULE.MAX_AMOUNT       
,MARGIN_COMMISSION_RULE.PERCENTAGE_AMOUNT
,MARGIN_COMMISSION_RULE.TYPE     
,MARGIN_COMMISSION_RULE.MATRIX_GROUP
FROM PAYMENT_METHODS_PACKAGE
LEFT JOIN MARGIN_COMMISSION_RULE ON PAYMENT_METHODS_PACKAGE.ID_CREDIT_CARD_COMMISSION=MARGIN_COMMISSION_RULE.ID
WHERE NOT ISNULL(PAYMENT_METHODS_PACKAGE.ID_CREDIT_CARD_COMMISSION); 

############## ESTRAE TUTTE LE CONFIGURAZIONI DI PAGAMENTO FORMATTATE PER CREARE IL FILE CSV PER IMPORT MASSIVO
SELECT PAYMENT_METHODS_PACKAGE.ID_PACKAGE
,PAYMENT_METHODS_PACKAGE.PAYMENT_METHOD
,REPLACE(MARGIN_COMMISSION_RULE.FLAT_AMOUNT,'.',',') AS 'FLAT_AMOUNT'     
,REPLACE(MARGIN_COMMISSION_RULE.MIN_AMOUNT,'.',',') AS 'MIN_AMOUNT'       
,REPLACE(MARGIN_COMMISSION_RULE.MAX_AMOUNT,'.',',') AS 'MAX_AMOUNT'        
,REPLACE(MARGIN_COMMISSION_RULE.PERCENTAGE_AMOUNT,'.',',') AS 'PERCENTAGE_AMOUNT' 
,MARGIN_COMMISSION_RULE.TYPE     
,IFNULL(MARGIN_COMMISSION_RULE.MATRIX_GROUP,'') AS 'MATRIX_GROUP'
,"0" AS 'DELETE'
FROM PAYMENT_METHODS_PACKAGE
LEFT JOIN MARGIN_COMMISSION_RULE ON PAYMENT_METHODS_PACKAGE.ID_CREDIT_CARD_COMMISSION=MARGIN_COMMISSION_RULE.ID
WHERE NOT ISNULL(PAYMENT_METHODS_PACKAGE.ID_CREDIT_CARD_COMMISSION); 



################################################################################################################
############## 				ESTRAE TUTTE LE CONFIGURAZIONI DI PAGAMENTO CHE HANNO ANCORA LA VECCHIA LOGICA ATTIVA	     ##############
################################################################################################################
select PAYMENT_METHODS_PACKAGE_PROFILE.*, PAYMENT_METHODS_PACKAGE.PAYMENT_METHOD, PAYMENT_METHODS_PACKAGE.COMMISSION, PAYMENT_METHODS_PACKAGE.COMMISSION_TYPE
from PAYMENT_METHODS_PACKAGE_PROFILE
LEFT JOIN PAYMENT_METHODS_PACKAGE ON PAYMENT_METHODS_PACKAGE_PROFILE.ID_PACKAGE=PAYMENT_METHODS_PACKAGE.ID_PACKAGE
WHERE PAYMENT_METHODS_PACKAGE_PROFILE.ID_PACKAGE IN
(SELECT ID_PACKAGE FROM PAYMENT_METHODS_PACKAGE WHERE ISNULL(ID_CREDIT_CARD_COMMISSION));


################################################################################################################
############## 				ESTRAE TUTTE LE CONFIGURAZIONI DI STAGING DIVERSE DA VOLAGRATIS							     ##############
################################################################################################################
SELECT *
FROM volagratis_staging.MARGIN_CONFIGURATION cob
WHERE
        not exists
	(
		select 1
		from volagratis.MARGIN_CONFIGURATION def
		where	
			cob.id_business_profile 		= def.id_business_profile and
			cob.id_airline_supplier_code		= def.id_airline_supplier_code and
			cob.route				= def.route and
			cob.id_global_commission		= def.id_global_commission and
			cob.id_margin_partition			= def.id_margin_partition and
			cob.id_discount				= def.id_discount 
	);
	
	
SELECT 'PRODUZIONE' as CONFIGURAZIONE, volagratis.MARGIN_CONFIGURATION.*
FROM volagratis.MARGIN_CONFIGURATION
WHERE ID_BUSINESS_PROFILE='VOLAGRATIS'
AND ID_AIRLINE_SUPPLIER_CODE='FR'
AND ROUTE='IT-IT'
UNION ALL
SELECT 'TEST' as CONFIGURAZIONE, volagratis_staging.MARGIN_CONFIGURATION.*
FROM volagratis_staging.MARGIN_CONFIGURATION
WHERE ID_BUSINESS_PROFILE='VOLAGRATIS'
AND ID_AIRLINE_SUPPLIER_CODE='FR'
AND ROUTE='IT-IT';


################################################################################################################
############## 	ESTRAE TUTTE LE CONFIGURAZIONI DI PRODUZIONE/STAGING IN BASE ALLA DATA CONFIGURAZIONE E CON LE NOTE       ##############
################################################################################################################
SELECT
ID_BUSINESS_PROFILE AS 'BUSINESS_PROFILE'
,ID_AIRLINE_SUPPLIER_CODE AS 'AIRLINE'
,ROUTE           
,REPLACE(FLAT_AMOUNT,'.',',') AS 'FLAT_AMOUNT'     
,REPLACE(MIN_AMOUNT,'.',',') AS 'MIN_AMOUNT'       
,REPLACE(MAX_AMOUNT,'.',',') AS 'MAX_AMOUNT'        
,REPLACE(PERCENTAGE_AMOUNT,'.',',') AS 'PERCENTAGE_AMOUNT'          
,IFNULL(MATRIX_GROUP,'') AS 'MARGIN_CORRECTION_MATRIX_GROUP'
,REPLACE(FLAT_DISCOUNT,'.',',') AS 'FLAT_DISCOUNT'
,REPLACE(MIN_DISCOUNT,'.',',') AS 'MIN_DISCOUNT'
,REPLACE(MAX_DISCOUNT,'.',',') AS 'MAX_DISCOUNT'
,REPLACE(PERCENTAGE_DISCOUNT,'.',',') AS 'PERCENTAGE_DISCOUNT'
,REPLACE(FLIGHT_COMMISSION,'.',',') AS 'FLIGHT_COMMISSION'
,REPLACE(AGENCY_COMMISSION,'.',',') AS 'AGENCY_COMMISSION'
,REPLACE(SUPPLEMENT_COMMISSION,'.',',') AS 'SUPPLEMENT_COMMISSION'
,ID_NOTE
,PARENT_NOTE
,NOTE
#,MODIFY_DATE
FROM MARGIN_CONFIGURATION
LEFT JOIN volagratis.MARGIN_COMMISSION_RULE ON MARGIN_CONFIGURATION.id_global_commission=volagratis.MARGIN_COMMISSION_RULE.id
LEFT JOIN volagratis.MARGIN_DISCOUNT_RULE ON MARGIN_CONFIGURATION.id_discount=volagratis.MARGIN_DISCOUNT_RULE.id
LEFT JOIN volagratis.MARGIN_PARTITION ON MARGIN_CONFIGURATION.id_margin_partition=volagratis.MARGIN_PARTITION.id
LEFT JOIN volagratis.MARGIN_NOTES ON MARGIN_CONFIGURATION.id_note=volagratis.MARGIN_NOTES.id
WHERE 
MODIFY_DATE >='2012-02-01'
#AND ID_BUSINESS_PROFILE='VOLAGRATIS'
ORDER BY MODIFY_DATE DESC;