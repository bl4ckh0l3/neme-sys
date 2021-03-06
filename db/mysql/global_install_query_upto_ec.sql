DROP TABLE IF EXISTS `attach_x_prodotti`;
CREATE TABLE IF NOT EXISTS `attach_x_prodotti` (  `id_prodotto` int(10) unsigned NOT NULL,  `id_attach` int(10) unsigned NOT NULL auto_increment,  `filename` varchar(100) NOT NULL,  `content_type` varchar(20) NOT NULL,  `path` varchar(100) NOT NULL,  `file_dida` text,  `file_label` varchar(2) NOT NULL,  PRIMARY KEY  (`id_attach`),  KEY `Index_2` (`id_prodotto`)) ENGINE=InnoDB  DEFAULT CHARSET=utf8 ;
DROP TABLE IF EXISTS `carrello`;
CREATE TABLE IF NOT EXISTS `carrello` (`id_carrello` int(10) unsigned NOT NULL auto_increment,  `id_utente` int(11) NOT NULL,  `dta_creazione` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,  PRIMARY KEY  (`id_carrello`)) ENGINE=InnoDB  DEFAULT CHARSET=utf8 ;
DROP TABLE IF EXISTS `ordini`;
CREATE TABLE IF NOT EXISTS `ordini` (  `id_ordine` int(10) unsigned NOT NULL auto_increment,  `id_utente` int(10) unsigned NOT NULL,  `dta_inserimento` timestamp NOT NULL default CURRENT_TIMESTAMP,  `stato_ordine` varchar(100) NOT NULL,  `totale_imponibile` DECIMAL(10,2) NOT NULL,  `totale_tasse` DECIMAL(10,2) NOT NULL,  `totale` decimal(10,2) NOT NULL,  `tipo_pagam` varchar(100), `payment_commission` decimal(10,2) NOT NULL default '0.00',  `pagam_effettuato` int(10) unsigned NOT NULL, `order_guid` varchar(250) NOT NULL, `user_notified_x_download` INT(1) UNSIGNED NOT NULL DEFAULT '0', `notes` text, `no_registration` smallint(1) unsigned NOT NULL DEFAULT '0', `id_ads` int(10) unsigned DEFAULT NULL, PRIMARY KEY  (`id_ordine`), KEY `Index_user` (`id_utente`)) ENGINE=InnoDB  DEFAULT CHARSET=utf8 ;
DROP TABLE IF EXISTS `prodotti`;
CREATE TABLE IF NOT EXISTS `prodotti` (  `id_prodotto` int(10) unsigned NOT NULL auto_increment,  `nome_prod` varchar(250) NOT NULL,  `sommario_prod` text,  `desc_prod` text,  `prezzo` decimal(10,2) NOT NULL,  `qta_disp` varchar(100) NOT NULL,  `attivo` int(10) unsigned NOT NULL,  `sconto` DECIMAL(10,2) NOT NULL default '0',  `codice_prod` varchar(100) NOT NULL,  `id_tassa_applicata` int(10) unsigned default NULL,  `prod_type` smallint(1) unsigned NOT NULL, `max_download` int(11) NOT NULL default '-1', `max_download_time` int(11) NOT NULL default '-1', `taxs_group` INT( 10 ) UNSIGNED DEFAULT NULL , `meta_description` TEXT default NULL,  `meta_keyword` TEXT default NULL,  `page_title` TEXT default NULL,  `edit_buy_qta` smallint(1) unsigned NOT NULL default '0',   PRIMARY KEY  (`id_prodotto`)) ENGINE=InnoDB  DEFAULT CHARSET=utf8 ;
DROP TABLE IF EXISTS `downloadable_products`;
CREATE TABLE IF NOT EXISTS `downloadable_products` ( `id` int(10) unsigned NOT NULL auto_increment,  `id_product` int(10) unsigned NOT NULL,  `filename` varchar(250) NOT NULL,  `path` varchar(250) NOT NULL,  `content_type` varchar(50) NOT NULL,  `file_size` int(10) unsigned NOT NULL,  `insert_date` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,  PRIMARY KEY  (`id`),  KEY `Index_2` (`id_product`),  KEY `Index_3` (`filename`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `down_prod_x_order`;
CREATE TABLE IF NOT EXISTS `down_prod_x_order` (`id` INT(11) UNSIGNED NOT NULL AUTO_INCREMENT,`id_order` INT(11) UNSIGNED NOT NULL ,`id_prod` INT(11) UNSIGNED NOT NULL ,`id_down_prod` INT(11) UNSIGNED NOT NULL ,`id_user` INT(11) UNSIGNED NOT NULL ,`active` SMALLINT(1) UNSIGNED NOT NULL default '0',`max_num_download` INT(3) NOT NULL  default '-1',`insert_date` TIMESTAMP NOT NULL ,`expire_date` TIMESTAMP NULL ,`download_counter` INT(3) UNSIGNED NOT NULL  default '0',`download_date` TIMESTAMP NULL, PRIMARY KEY  (`id`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `prodotti_x_carrello`;
CREATE TABLE IF NOT EXISTS `prodotti_x_carrello` (  `id_carrello` int(10) unsigned NOT NULL,  `id_prodotto` int(10) unsigned NOT NULL,  `counter_prod` int(10) unsigned NOT NULL,  `qta_prod` int(10) unsigned NOT NULL, `prod_type` smallint(1) unsigned NOT NULL, PRIMARY KEY  (`id_carrello`,`id_prodotto`,`counter_prod`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `prodotti_x_ordine`;
CREATE TABLE IF NOT EXISTS `prodotti_x_ordine` ( `id_ordine` int(10) unsigned NOT NULL, `id_prodotto` int(10) unsigned NOT NULL, `counter_prod` int(10) unsigned NOT NULL, `nome_prodotto` varchar(100) NOT NULL,  `qta` int(10) unsigned NOT NULL,  `totale` decimal(10,2) NOT NULL,  `tax` DECIMAL(10,2) NOT NULL default '0.00',  `desc_tax` varchar(100) DEFAULT NULL,   `prod_type` smallint(1) unsigned NOT NULL default '0',  PRIMARY KEY  (`id_ordine`,`id_prodotto`,`counter_prod`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `relation_x_prodotto`;
CREATE TABLE IF NOT EXISTS `relation_x_prodotto` (  `id_prod` int(10) unsigned NOT NULL,  `id_prod_rel` int(10) unsigned NOT NULL,  UNIQUE KEY `Index_Rp` (`id_prod`,`id_prod_rel`),  INDEX `Index_RpP`(`id_prod`),  INDEX `Index_RpPr`(`id_prod_rel`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `target_x_prodotto`;
CREATE TABLE IF NOT EXISTS `target_x_prodotto` (  `id_target` int(10) unsigned NOT NULL,  `id_prodotto` int(10) unsigned NOT NULL,  KEY `Index_1` (`id_prodotto`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `spese_accessorie`;
CREATE TABLE IF NOT EXISTS `spese_accessorie` (  `id` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,  `descrizione_spesa` VARCHAR(100) NOT NULL, `valore` DECIMAL(10,2) NOT NULL,  `tipologia_valore` SMALLINT(1) UNSIGNED NOT NULL,  `id_tassa_applicata` INTEGER(10) UNSIGNED default NULL,  `applica_frontend`SMALLINT(1) UNSIGNED,  `applica_backend` SMALLINT(1) UNSIGNED,  `autoactive` SMALLINT(1) UNSIGNED NOT NULL default '0',  `multiply` SMALLINT(1) UNSIGNED NOT NULL default '0', `required` SMALLINT(1) UNSIGNED NOT NULL default '0',  `group` VARCHAR(50) NOT NULL, `taxs_group` INT( 10 ) UNSIGNED DEFAULT NULL ,  `type_view` SMALLINT(1) UNSIGNED NOT NULL default '0',  PRIMARY KEY (`id`),  INDEX `Index_2`(`valore`),  INDEX `Index_3`(`id_tassa_applicata`)) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;
DROP TABLE IF EXISTS `tasse`;
CREATE TABLE IF NOT EXISTS `tasse` (  `id` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,  `descrizione_tassa` VARCHAR(100) NOT NULL,  `valore` DECIMAL(10,2) NOT NULL,  `tipologia_valore` SMALLINT(1) UNSIGNED NOT NULL,  PRIMARY KEY (`id`),  INDEX `Index_2`(`valore`))ENGINE=InnoDB  DEFAULT CHARSET=utf8 ;
DROP TABLE IF EXISTS `payment_type`;
CREATE TABLE IF NOT EXISTS `payment_type` (  `id` int(10) unsigned NOT NULL auto_increment,  `keyword_multilingua` varchar(250) default NULL,  `descrizione` varchar(250) default NULL,  `dati_pagamento` varchar(250) NOT NULL, `commission` decimal(10,2) NOT NULL default '0.00', `commission_type` SMALLINT(1) UNSIGNED NOT NULL  default '1', `url` smallint(1) unsigned NOT NULL default '0',  `id_modulo` int(10) default NULL,  `activate` smallint(1) unsigned NOT NULL default '0',  `payment_type` smallint(1) unsigned NOT NULL default '0', PRIMARY KEY  (`id`)) ENGINE=InnoDB  DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `payment_field`;
CREATE TABLE IF NOT EXISTS `payment_field` (  `id` int(10) unsigned NOT NULL auto_increment,  `id_payment` int(10) unsigned NOT NULL,  `id_modulo` int(10) unsigned default NULL,  `name` varchar(50) NOT NULL,  `value` varchar(250) default NULL,  `match_field` varchar(50) default NULL,  PRIMARY KEY  USING BTREE (`id`),  UNIQUE KEY `Index_UX` (`id_payment`,`id_modulo`,`name`,`match_field`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `payment_modulo`;
CREATE TABLE IF NOT EXISTS `payment_modulo` (  `id` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,  `name` VARCHAR(45) NOT NULL,  `directory` VARCHAR(100) NOT NULL,  `logo` TEXT,  `insert_page` VARCHAR(100) NOT NULL,  `checkout_page` VARCHAR(100) NOT NULL,  `checkin_page` VARCHAR(100) NOT NULL, `checkin_fault_page` VARCHAR(100) NOT NULL, `id_ordine_field` VARCHAR(100) NOT NULL,  `ip_provider` VARCHAR(150) NOT NULL, PRIMARY KEY (`id`),  INDEX `Index_2`(`name`))ENGINE = InnoDB  DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `payment_fixed_app_field`;
CREATE TABLE IF NOT EXISTS `payment_fixed_app_field` (  `id` int(10) unsigned NOT NULL auto_increment,  `keyword` varchar(50) NOT NULL,  `value` varchar(100) default NULL,  `used` smallint(1) unsigned NOT NULL default '1',  PRIMARY KEY  (`id`),  UNIQUE KEY `Index_2` (`keyword`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `paypal_field`;
CREATE TABLE IF NOT EXISTS `paypal_field` (  `id` int(10) unsigned NOT NULL auto_increment,  `keyword` varchar(50) NOT NULL,  `value` varchar(100) default NULL,  `match_field` varchar(50) default NULL,  PRIMARY KEY  (`id`),  UNIQUE KEY `Index_2` (`keyword`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `sella_field`;
CREATE TABLE `sella_field` (  `id` int(10) unsigned NOT NULL auto_increment,  `keyword` varchar(50) NOT NULL,  `value` varchar(100) default NULL,  `match_field` varchar(50) default NULL,  PRIMARY KEY  (`id`),  UNIQUE KEY `Index_2` (`keyword`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `spese_x_ordine`;
CREATE TABLE IF NOT EXISTS `spese_x_ordine` (  `id_ordine` INTEGER UNSIGNED NOT NULL,  `id_spesa` INTEGER UNSIGNED NOT NULL,  `imponibile` DECIMAL(10,2) NOT NULL,  `tasse` DECIMAL(10,2) NOT NULL,  `totale` DECIMAL(10,2) NOT NULL, `desc_spesa` varchar(100) DEFAULT NULL,  INDEX `Index_1`(`id_ordine`),  INDEX `Index_2`(`id_spesa`))ENGINE = InnoDB  DEFAULT CHARSET=utf8 ;
DDROP TABLE IF EXISTS `payment_transactions`;
CREATE TABLE IF NOT EXISTS `payment_transactions` (  `id` int(11) unsigned NOT NULL auto_increment,  `id_ordine` int(11) unsigned NOT NULL, `id_modulo` INTEGER UNSIGNED NOT NULL,  `id_transaction` varchar(100) NOT NULL,  `status` varchar(50) default NULL,  `notified` smallint(1) unsigned NOT NULL default '0', `insert_date` datetime NOT NULL,  PRIMARY KEY  (`id`),  INDEX `Index_2` (`id_ordine`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `currency`;
CREATE TABLE IF NOT EXISTS `currency` ( `id` int(10) unsigned NOT NULL auto_increment,  `currency` varchar(5) NOT NULL,  `rate` decimal(10,4) NOT NULL, `dta_riferimento` date NOT NULL,  `dta_inserimento` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,  `active` int(1) unsigned NOT NULL default '0',  `is_default` int(1) unsigned NOT NULL default '0',  PRIMARY KEY  (`id`),  KEY `currency` (`currency`)) ENGINE=InnoDB  DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `shipping_address`;
CREATE TABLE IF NOT EXISTS `shipping_address` (  `id` int(10) unsigned NOT NULL auto_increment,  `id_user` int(10) unsigned NOT NULL,  `name` varchar(100) default NULL,  `surname` varchar(100) default NULL,  `cfiscvat` varchar(16) default NULL,  `address` varchar(250) default NULL,  `city` varchar(100) default NULL,  `zipCode` varchar(20) default NULL,  `country` varchar(100) default NULL,  `state_region` varchar(100) default NULL, `is_company_client` SMALLINT(1) UNSIGNED NOT NULL default '0', PRIMARY KEY  (`id`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `order_shipping_address`;
CREATE TABLE IF NOT EXISTS `order_shipping_address` ( `id_order` int(10) unsigned NOT NULL,  `id_shipping` int(10) unsigned NOT NULL,  `address` varchar(250) default NULL,  `city` varchar(100) default NULL,  `zipCode` varchar(20) default NULL,  `country` varchar(100) default NULL,  `state_region` varchar(100) default NULL, `is_company_client` SMALLINT(1) UNSIGNED NOT NULL default '0') ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `bills_address`;
CREATE TABLE IF NOT EXISTS `bills_address` (  `id` int(10) unsigned NOT NULL auto_increment,  `id_user` int(10) unsigned NOT NULL,  `name` varchar(100) default NULL,  `surname` varchar(100) default NULL,  `cfiscvat` varchar(16) default NULL,  `address` varchar(250) default NULL,  `city` varchar(100) default NULL,  `zipCode` varchar(20) default NULL,  `country` varchar(100) default NULL,  `state_region` varchar(100) default NULL,  PRIMARY KEY  (`id`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `order_bills_address`;
CREATE TABLE IF NOT EXISTS `order_bills_address` (  `id_order` int(10) unsigned NOT NULL,  `id_bills` int(10) unsigned NOT NULL,  `address` varchar(250) default NULL,  `city` varchar(100) default NULL,  `zipCode` varchar(20) default NULL,  `country` varchar(100) default NULL,  `state_region` varchar(100) default NULL) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `user_group`;
CREATE TABLE IF NOT EXISTS `user_group` (  `id` int(11) unsigned NOT NULL auto_increment,  `short_desc` varchar(100) NOT NULL,  `long_desc` text, `default` int(1) unsigned NOT NULL default '0', `taxs_group` INT( 10 ) UNSIGNED DEFAULT NULL ,   PRIMARY KEY  (`id`)) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;
DROP TABLE IF EXISTS `margin_discount`;
CREATE TABLE IF NOT EXISTS `margin_discount` (  `id` int(11) unsigned NOT NULL auto_increment,  `margin` decimal(10,2) unsigned NOT NULL,  `discount` decimal(10,2) unsigned NOT NULL,  `apply_prod_discount` int(1) unsigned NOT NULL default '0',  `apply_user_discount` int(1) unsigned NOT NULL default '0', PRIMARY KEY  (`id`)) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;
DROP TABLE IF EXISTS `usr_group_x_margin_disc`;
CREATE TABLE IF NOT EXISTS `usr_group_x_margin_disc` (  `id_marg_disc` int(11) unsigned NOT NULL,  `id_user_group` int(11) unsigned NOT NULL,  KEY `Index_1` (`id_user_group`),  UNIQUE KEY `Index_2` (`id_user_group`),  UNIQUE KEY `Index_3` (`id_user_group`,`id_marg_disc`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `product_fields`;
CREATE TABLE `product_fields` (  `id` int(11) unsigned NOT NULL auto_increment,  `description` varchar(100) NOT NULL,  `id_group` int(11) unsigned DEFAULT NULL,  `type` int(11) unsigned NOT NULL,  `type_content` int(11) unsigned NOT NULL,  `order` int(3) unsigned NOT NULL DEFAULT 0,  `required` int(1) UNSIGNED NOT NULL DEFAULT 0,  `enabled` int(1) UNSIGNED NOT NULL DEFAULT 0,  `max_lenght` int(3) UNSIGNED DEFAULT NULL,  `editable` int(1) UNSIGNED NOT NULL DEFAULT 0,  PRIMARY KEY  (`id`),  KEY `Index_3` (`id_group`),  KEY `Index_4` (`type`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `product_fields_values`;
CREATE TABLE `product_fields_values` (  `id_field` int(11) unsigned NOT NULL,  `value` varchar(250) NOT NULL,  `order` int(3) unsigned NOT NULL DEFAULT 0,  UNIQUE KEY `Index_PFV` (`id_field`,`value`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `product_fields_group`;
CREATE TABLE `product_fields_group` (  `id` int(11) unsigned NOT NULL auto_increment,  `description` varchar(100) NOT NULL,  `order` int(2) unsigned NOT NULL DEFAULT 0,  PRIMARY KEY  (`id`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `product_fields_type`;
CREATE TABLE `product_fields_type` (  `id` int(11) unsigned NOT NULL auto_increment,  `description` varchar(100) NOT NULL,  PRIMARY KEY  (`id`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `product_fields_type_content`;
CREATE TABLE `product_fields_type_content` (  `id` int(11) unsigned NOT NULL auto_increment,  `description` varchar(100) NOT NULL,  PRIMARY KEY  (`id`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `product_fields_match`;
CREATE TABLE `product_fields_match` (  `id_field` INTEGER UNSIGNED NOT NULL,  `id_prod` INTEGER UNSIGNED NOT NULL,  `value` varchar(250) NOT NULL,  PRIMARY KEY (`id_field`, `id_prod`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `product_fields_value_match`;
CREATE TABLE `product_fields_value_match` (  `id_field` INTEGER UNSIGNED NOT NULL,  `id_prod` INTEGER UNSIGNED NOT NULL,  `qta_prod` int(10) NOT NULL,  `value` varchar(250) NOT NULL,  PRIMARY KEY (`id_field`, `id_prod`, `value`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `product_fields_rel_value_match`;
CREATE TABLE IF NOT EXISTS `product_fields_rel_value_match` (  `id_prod` int(10) unsigned NOT NULL,  `id_field` int(10) unsigned NOT NULL,  `field_val` varchar(250) NOT NULL,  `id_field_rel` int(10) unsigned NOT NULL,  `field_rel_val` varchar(250) NOT NULL,  `qta_rel` int(10) NOT NULL,  KEY `id_prod` (`id_prod`),  KEY `id_field` (`id_field`),  KEY `field_val` (`field_val`),  KEY `id_field_rel` (`id_field_rel`),  KEY `field_rel_val` (`field_rel_val`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `product_fields_x_order`;
CREATE TABLE IF NOT EXISTS `product_fields_x_order` (  `counter` INTEGER UNSIGNED NOT NULL,  `id_order` int(10) unsigned NOT NULL,  `id_prod` int(10) unsigned NOT NULL,  `id_field` int(10) unsigned NOT NULL,  `qta_prod` int(10) unsigned NOT NULL,  `value` varchar(250) NOT NULL,  PRIMARY KEY  (`counter`,`id_order`,`id_prod`,`id_field`,`value`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `product_fields_x_card`;
CREATE TABLE IF NOT EXISTS `product_fields_x_card` (  `counter` INTEGER UNSIGNED NOT NULL,  `id_card` int(10) unsigned NOT NULL,  `id_prod` int(10) unsigned NOT NULL,  `id_field` int(10) unsigned NOT NULL,  `qta_prod` int(10) unsigned NOT NULL,  `value` varchar(250) NOT NULL,  PRIMARY KEY  (`counter`,`id_card`,`id_prod`,`id_field`,`value`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `tax_group`;
CREATE TABLE IF NOT EXISTS `tax_group` ( `id` int(10) unsigned NOT NULL auto_increment,  `description` VARCHAR(100) NOT NULL,  PRIMARY KEY (`id`),  INDEX `Index_TG_dc`(`description`)) ENGINE = InnoDB  DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `tax_group_value`;
CREATE TABLE IF NOT EXISTS `tax_group_value` ( `id_group` int(10) unsigned NOT NULL,  `country_code` VARCHAR(2) NOT NULL,  `state_region_code` VARCHAR(10) DEFAULT NULL,  `id_tassa_applicata` int(10) unsigned default NULL, `exclude_calculation` SMALLINT(1) UNSIGNED NOT NULL default '0', INDEX `Index_TGV_ig`(`id_group`),  INDEX `Index_TGV_cc`(`country_code`), INDEX `Index_TGV_src`(`state_region_code`)) ENGINE = InnoDB  DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `prodotto_main_field_translation`;
CREATE TABLE IF NOT EXISTS `prodotto_main_field_translation` (  `id_prod` int(10) unsigned NOT NULL,  `main_field` int(3) unsigned NOT NULL,  `lang_code` varchar(2) NOT NULL,  `value` text,  UNIQUE KEY `Index_Pmft` (`id_prod`,`main_field`,`lang_code`),  INDEX `Index_Pmfti`(`id_prod`),  INDEX `Index_Pmftm`(`main_field`),  INDEX `Index_Pmftl`(`lang_code`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `ads_promotion`;
CREATE TABLE IF NOT EXISTS `ads_promotion` (  `id_ads` int(10) unsigned NOT NULL,  `id_element` int(10) unsigned NOT NULL,  `cod_element` VARCHAR(100) NOT NULL,  `active` SMALLINT(1) UNSIGNED NOT NULL default '0',  `dta_inserimento` timestamp NOT NULL default CURRENT_TIMESTAMP,  PRIMARY KEY  (`id_ads`,`id_element`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `business_rules`;
CREATE TABLE IF NOT EXISTS `business_rules` (`id` int(10) unsigned NOT NULL auto_increment,  `rule_type` int(10) unsigned NOT NULL,  `label` varchar(100) NOT NULL,  `description` text,  `activate` smallint(1) unsigned NOT NULL default '0',  `voucher_id` int(10) default NULL,  PRIMARY KEY  (`id`),  KEY `Index_label` (`label`)) ENGINE=InnoDB  DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `business_rules_config`;
CREATE TABLE IF NOT EXISTS `business_rules_config` (  `id` int(10) unsigned NOT NULL auto_increment,  `id_rule` int(10) unsigned NOT NULL,  `id_prod_orig` int(10) unsigned default NULL,  `id_prod_ref` int(10) unsigned default NULL,  `rate_from` decimal(10,2) NOT NULL,  `rate_to` decimal(10,2) NOT NULL,  `rate_from_ref` decimal(10,2) default NULL,  `rate_to_ref` decimal(10,2) default NULL,  `operation` smallint(1) unsigned NOT NULL default '0' COMMENT 'tipo di operatione da eseguire: 0 nulla, 1 somma, 2 sottrazione;',  `applyto` smallint(1) unsigned NOT NULL default '0' COMMENT 'a chi deve essere applicata la regola: 0 nulla, 1 a prod_orig, 2 a prod_ref, 3 al meno caro, 4 al più caro, 5 a tutti e due;',  `apply_4_qta` int(10) unsigned NOT NULL default '0' COMMENT 'per quale quantità  deve essere applicata la regola',  `valore` decimal(10,2) NOT NULL,  PRIMARY KEY  (`id`),  UNIQUE KEY `Index_U` (`id_rule`,`id_prod_orig`,`rate_from`,`rate_to`,`id_prod_ref`),  KEY `Index_From` (`rate_from`),  KEY `Index_To` (`rate_to`),  KEY `Index_Val` (`valore`)) ENGINE=InnoDB  DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `business_rules_x_ordine`;
CREATE TABLE IF NOT EXISTS `business_rules_x_ordine` (  `id_rule` int(10) unsigned NOT NULL,  `id_order` int(10) unsigned NOT NULL,  `id_prod` int(10) unsigned NOT NULL default '0', `counter_prod` int(10) unsigned NOT NULL default '0', `label` VARCHAR(100) NOT NULL,  `valore` DECIMAL(10,2) NOT NULL,     PRIMARY KEY (`id_rule`,`id_order`,`id_prod`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `voucher_campaign`;
CREATE TABLE IF NOT EXISTS `voucher_campaign` (  `id` int(10) unsigned NOT NULL auto_increment,  `label` varchar(100) NOT NULL,  `voucher_type` int(10) unsigned NOT NULL default '0' COMMENT 'tipo di voucher creato - 0=one shot, 1=per x volte, 2=one shot entro il periodo specificato, 3=per x volte entro il periodo specificato, 4=gift (come one shot ma con id_utente che ha fatto il regalo associato al voucher)',  `description` text,  `valore` decimal(10,2) NOT NULL,  `operation` smallint(1) unsigned NOT NULL default '0' COMMENT 'tipo di calcolo applicato - 0=percentuale, 1=fisso',  `activate` smallint(1) unsigned NOT NULL default '0',  `max_generation` int(10) NOT NULL default '-1',  `max_usage` int(10) NOT NULL default '-1',  `enable_date` timestamp NULL default NULL,  `expire_date` timestamp NULL default NULL,  `exclude_prod_rule` smallint(1) unsigned NOT NULL default '0',  PRIMARY KEY  (`id`),  KEY `Index_label` (`label`)) ENGINE=InnoDB  DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `voucher_code`;
CREATE TABLE IF NOT EXISTS `voucher_code` (  `id` int(10) unsigned NOT NULL auto_increment,  `code` varchar(100) NOT NULL COMMENT 'il codice voucher verrà generato con un nuovo GUID ad hoc',  `voucher_campaign` int(10) unsigned NOT NULL,  `insert_date` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,  `usage_counter` int(10) unsigned NOT NULL default '0',  `id_user_ref` int(10) unsigned default NULL,  PRIMARY KEY  (`id`,`code`),  KEY `Index_vc` (`voucher_campaign`)) ENGINE=InnoDB  DEFAULT CHARSET=utf8;
DROP TABLE IF EXISTS `voucher_x_ordine`;
CREATE TABLE IF NOT EXISTS `voucher_x_ordine` (  `id_order` int(10) unsigned NOT NULL,  `voucher_code` varchar(100) NOT NULL,  `id_voucher` int(10) unsigned NOT NULL,  `valore` decimal(10,2) NOT NULL, `insert_date` TIMESTAMP NOT NULL , PRIMARY KEY  (`id_order`,`voucher_code`,`id_voucher`, `insert_date`)) ENGINE=InnoDB DEFAULT CHARSET=utf8;
ALTER TABLE `attach_x_prodotti` ADD FOREIGN KEY (`id_prodotto`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `prodotti_x_carrello` ADD FOREIGN KEY (`id_carrello`) REFERENCES `carrello` (`id_carrello`) ON DELETE CASCADE;
ALTER TABLE `prodotti_x_ordine` ADD FOREIGN KEY (`id_ordine`) REFERENCES `ordini` (`id_ordine`) ON DELETE CASCADE;
ALTER TABLE `target_x_prodotto` ADD FOREIGN KEY (`id_prodotto`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `spese_x_ordine` ADD FOREIGN KEY (`id_ordine`) REFERENCES `ordini` (`id_ordine`) ON DELETE CASCADE;
ALTER TABLE `shipping_address` ADD FOREIGN KEY (`id_user`) REFERENCES `utenti` (`id`) ON DELETE CASCADE;
ALTER TABLE `order_shipping_address` ADD FOREIGN KEY (`id_order`) REFERENCES `ordini` (`id_ordine`) ON DELETE CASCADE;
ALTER TABLE `order_shipping_address` ADD FOREIGN KEY (`id_shipping`) REFERENCES `shipping_address` (`id`) ON DELETE CASCADE;
ALTER TABLE `downloadable_products` ADD FOREIGN KEY (`id_product`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `usr_group_x_margin_disc` ADD FOREIGN KEY (`id_marg_disc`) REFERENCES `margin_discount` (`id`) ON DELETE CASCADE;
ALTER TABLE `usr_group_x_margin_disc` ADD FOREIGN KEY (`id_user_group`) REFERENCES `user_group` (`id`) ON DELETE CASCADE;
ALTER TABLE `product_fields_values` ADD FOREIGN KEY (`id_field`) REFERENCES `product_fields` (`id`) ON DELETE CASCADE;
ALTER TABLE `product_fields_match` ADD FOREIGN KEY (`id_prod`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `product_fields_match` ADD FOREIGN KEY (`id_field`) REFERENCES `product_fields` (`id`) ON DELETE CASCADE;
ALTER TABLE `product_fields_value_match` ADD FOREIGN KEY (`id_prod`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `product_fields_value_match` ADD FOREIGN KEY (`id_field`) REFERENCES `product_fields` (`id`) ON DELETE CASCADE;
ALTER TABLE `product_fields_x_order` ADD FOREIGN KEY (`id_order`) REFERENCES `ordini` (`id_ordine`) ON DELETE CASCADE;
ALTER TABLE `product_fields` ADD FOREIGN KEY (`id_group`) REFERENCES `product_fields_group` (`id`);
ALTER TABLE `product_fields` ADD FOREIGN KEY (`type`) REFERENCES `product_fields_type` (`id`);
ALTER TABLE `product_fields_x_card` ADD FOREIGN KEY (`id_card`) REFERENCES `carrello` (`id_carrello`) ON DELETE CASCADE;
ALTER TABLE `prodotti_x_carrello` ADD FOREIGN KEY (`id_carrello`) REFERENCES `carrello` (`id_carrello`) ON DELETE CASCADE;
ALTER TABLE `relation_x_prodotto` ADD FOREIGN KEY (`id_prod`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `relation_x_prodotto` ADD FOREIGN KEY (`id_prod_rel`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `tax_group_value` ADD FOREIGN KEY (`id_group`) REFERENCES `tax_group` (`id`) ON DELETE CASCADE;
ALTER TABLE `prodotto_main_field_translation` ADD FOREIGN KEY (`id_prod`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `ads_promotion` ADD FOREIGN KEY (`id_ads`) REFERENCES `ads` (`id_ads`) ON DELETE CASCADE;
ALTER TABLE `ads_promotion` ADD FOREIGN KEY (`id_element`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `spese_accessorie_config` ADD FOREIGN KEY (`id_spesa`) REFERENCES `spese_accessorie` (`id`) ON DELETE CASCADE;
ALTER TABLE `business_rules_config` ADD FOREIGN KEY (`id_rule`) REFERENCES `business_rules` (`id`) ON DELETE CASCADE;
ALTER TABLE `business_rules_x_ordine` ADD FOREIGN KEY (`id_order`) REFERENCES `ordini` (`id_ordine`) ON DELETE CASCADE;
ALTER TABLE `voucher_x_ordine` ADD FOREIGN KEY (`id_order`) REFERENCES `ordini` (`id_ordine`) ON DELETE CASCADE;
DELETE FROM `module_portal` WHERE `keyword` = 'neme-sys';
DELETE FROM `module_portal` WHERE `keyword` = 'econeme-sys';
INSERT INTO `module_portal` (`keyword`, `descrizione`, `version`, `active`) VALUES ('econemesys', 'base econeme-sys installation', 'buildVersSql', '1');
INSERT INTO `target_type` (`id`, `descrizione`) VALUES (2, 'backend.target.detail.table.select.option.target_type_prod');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('day_carrello_is_valid', 'backend.config.lista.table.description.day_carrello_is_valid', '7', '0', 'ecom_card');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('del_carrello_on_exit', 'backend.config.lista.table.description.del_carrello_on_exit', '0', '0', 'ecom_card');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('dir_upload_prod', 'backend.config.lista.table.description.dir_upload_prod', '/public/upload/files/prod/', '1', 'directory');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('dir_down_prod', 'backend.config.lista.table.description.dir_down_prod', '/app_data/', '1', 'directory');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('mail_order_receiver', 'backend.config.lista.table.description.mail_order_receiver', '', '0', 'mail_order');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('mail_order_sender', 'backend.config.lista.table.description.mail_order_sender', '', '0', 'mail_order');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('mail_order_bcc', 'backend.config.lista.table.description.mail_order_bcc', '', '0', 'mail_order');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('mail_order_cc', 'backend.config.lista.table.description.mail_order_cc', '', '0', 'mail_order');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('minute_order_modify_permit', 'backend.config.lista.table.description.minute_order_modify_permit', '1440', '0', 'ecom_time');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('manage_sconti', 'backend.config.lista.table.description.manage_sconti', '0', '0', 'ecom_sconti');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('disable_ecommerce', 'backend.config.lista.table.description.disable_ecommerce', '0', '0', 'ecom_disable');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('show_ship_box', 'backend.config.lista.table.description.show_ship_box', '1', '0', 'ecom_info');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('show_bills_box', 'backend.config.lista.table.description.show_bills_box', '1', '0', 'ecom_info');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('show_user_field_on_direct_buy', 'backend.config.lista.table.description.show_user_field_on_direct_buy', '0', '0', 'ecom_info');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('enable_international_tax_option', 'backend.config.lista.table.description.enable_international_tax_option', '0', '1', 'ecom_tax');
INSERT INTO `config_portal` (`keyword`, `descrizione`, `conf_value`, `alert`, `tipo`) VALUES ('enable_ads', 'backend.config.lista.table.description.enable_ads', '0', '0', 'configuration_ads');
INSERT INTO `payment_fixed_app_field` (`id`,`keyword`,`value`,`used`) VALUES (1,'id_order_ack','ID Ordine',1);
INSERT INTO `payment_fixed_app_field` (`id`,`keyword`,`value`,`used`) VALUES (2,'amount_order_ack','Totale Ordine',1);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (1,'custom',NULL,'id_order_ack');
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (2,'amount',NULL,'amount_order_ack');
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (3,'external_url','https://www.paypal.com/cgi-bin/webscr',NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (4,'cmd','_xclick',NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (5,'business',NULL,NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (6,'ack','Success',NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (7,'return','http://localhost/common/include/checkin.asp',NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (8,'cancel_return','http://localhost/common/include/checkin_fault.asp',NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (9,'item_name',NULL,NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (10,'currency_code','EUR',NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (11,'image_url',NULL,NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (12,'notify_url','http://localhost/editor/payments/moduli/paypal/checkin_notify.asp',NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (13,'tx',NULL,NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (14,'at',NULL,NULL);
INSERT INTO `paypal_field` (`id`,`keyword`,`value`,`match_field`) VALUES (15,'cm',NULL,'id_order_ack');
INSERT INTO `sella_field` (`id`,`keyword`,`value`,`match_field`) VALUES (1,'shoptransactionid',NULL,'id_order_ack');
INSERT INTO `sella_field` (`id`,`keyword`,`value`,`match_field`) VALUES (2,'amount',NULL,'amount_order_ack');
INSERT INTO `sella_field` (`id`,`keyword`,`value`,`match_field`) VALUES (3,'external_url','https://testecomm.sella.it/gestpay/pagam.asp',NULL);
INSERT INTO `sella_field` (`id`,`keyword`,`value`,`match_field`) VALUES (4,'currency','242',NULL);
INSERT INTO `sella_field` (`id`,`keyword`,`value`,`match_field`) VALUES (5,'shoplogin','GESPAY47944',NULL);
INSERT INTO `sella_field` (`id`,`keyword`,`value`,`match_field`) VALUES (6,'a','GESPAY47944',NULL);
INSERT INTO `sella_field` (`id`,`keyword`,`value`,`match_field`) VALUES (7,'b',NULL,'id_order_ack');
INSERT INTO `payment_modulo` (`id`,`name`,`directory`,`logo`,`insert_page`,`checkout_page`,`checkin_page`,`checkin_fault_page`,`id_ordine_field`,`ip_provider`) VALUES (1,'paypal','paypal','<a href="#" onclick=javascript:window.open("https://www.paypal.com/us/cgi-bin/webscr?cmd=xpt/Marketing/popup/OLCWhatIsPayPal-outside","olcwhatispaypal","toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes, width=400, height=350");><img  src="https://www.paypal.com/en_US/i/logo/PayPal_mark_50x34.gif" border="0" alt="Acceptance Mark"></a>','insert.asp','checkout.asp','checkin.asp','checkin_fault.asp','custom|cm','212.48.8.140|87.0.139.170');
INSERT INTO `payment_modulo` (`id`,`name`,`directory`,`logo`,`insert_page`,`checkout_page`,`checkin_page`,`checkin_fault_page`,`id_ordine_field`,`ip_provider`) VALUES (2,'sella','sella',NULL,'insert.asp','checkout.asp','checkin.asp','checkin_fault.asp','shoptransactionid|b','');
INSERT INTO `currency` (`id`,`currency`,`rate`,`dta_riferimento`,`dta_inserimento`,`active`,`is_default`) VALUES(1,'EUR','1.0000','0000-00-00','0000-00-00 00:00:00',1,1);
INSERT INTO `product_fields_type` (`id`, `description`) VALUES (1, 'text');
INSERT INTO `product_fields_type` (`id`, `description`) VALUES (2, 'textarea');
INSERT INTO `product_fields_type` (`id`, `description`) VALUES (3, 'select');
INSERT INTO `product_fields_type` (`id`, `description`) VALUES (4, 'select-multiple');
INSERT INTO `product_fields_type` (`id`, `description`) VALUES (5, 'checkbox');
INSERT INTO `product_fields_type` (`id`, `description`) VALUES (6, 'radio');
INSERT INTO `product_fields_type` (`id`, `description`) VALUES (7, 'hidden');
INSERT INTO `product_fields_type` (`id`, `description`) VALUES (8, 'file');
INSERT INTO `product_fields_type` (`id`, `description`) VALUES (9, 'editor-html');
INSERT INTO `product_fields_type_content` (`id`, `description`) VALUES (1, 'text');
INSERT INTO `product_fields_type_content` (`id`, `description`) VALUES (2, 'integer');
INSERT INTO `product_fields_type_content` (`id`, `description`) VALUES (3, 'decimal');
INSERT INTO `product_fields_type_content` (`id`, `description`) VALUES (4, 'date');