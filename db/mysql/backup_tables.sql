-- --------------------------------------------------------
-- 
-- Struttura della tabella `categorie`
-- 
DROP TABLE IF EXISTS `categorie`;
CREATE TABLE IF NOT EXISTS `categorie` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `num_menu` smallint(2) NOT NULL default '1',
  `gerarchia` varchar(100) NOT NULL,
  `descrizione` varchar(250) default NULL,
  `type` varchar(100) NOT NULL,
  `contiene_news` int(1) unsigned NOT NULL default '0',
  `contiene_prod` int(1) unsigned NOT NULL default '0',
  `visibile` int(1) unsigned NOT NULL default '0',
  `id_template` int(10) default NULL,
  `meta_description` TEXT default NULL,
  `meta_keyword` TEXT default NULL,
  `page_title` TEXT default NULL,
  `sub_domain_url` varchar(250) default NULL,  
  PRIMARY KEY  (`id`),
  KEY `num_menu` (`num_menu`),
  KEY `gerarchia` (`gerarchia`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `commenti`
-- 
DROP TABLE IF EXISTS `commenti`;
CREATE TABLE IF NOT EXISTS `commenti` (
  `id_commento` int(10) unsigned NOT NULL auto_increment,
  `id_element` int(10) unsigned NOT NULL,
  `element_type` int(3) unsigned NOT NULL,
  `id_utente` int(10) NOT NULL,
  `message` text,
  `dta_inserimento` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `vote_type` int(1) unsigned NOT NULL default '0',
  `active` int(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (`id_commento`),
  KEY `Index_2` (`id_element`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `config_portal`
-- 
DROP TABLE IF EXISTS `config_portal`;
CREATE TABLE IF NOT EXISTS `config_portal` (
  `keyword` varchar(100) NOT NULL,
  `descrizione` varchar(250) default NULL,
  `conf_value` TEXT default NULL,
  `alert` char(1) NOT NULL default '0',
  `tipo` varchar(100) NOT NULL,
  PRIMARY KEY  (`keyword`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `module_portal`
-- 
DROP TABLE IF EXISTS `module_portal`;
CREATE TABLE IF NOT EXISTS `module_portal` (
  `keyword` varchar(100) NOT NULL,
  `descrizione` TEXT default NULL,
  `version` varchar(100) NOT NULL,
  `active` char(1) NOT NULL default '0',
  `date_insert` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`keyword`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `conferma_utente`
-- 
DROP TABLE IF EXISTS `conferma_utente`;
CREATE TABLE IF NOT EXISTS `conferma_utente` (
  `id_user` int(10) unsigned NOT NULL,
  `confirmation_code` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `file_x_news`
-- 
DROP TABLE IF EXISTS `file_x_news`;
CREATE TABLE IF NOT EXISTS `file_x_news` (
  `id_news` int(10) unsigned NOT NULL,
  `id_file` int(10) unsigned NOT NULL,
  KEY `Index_1` (`id_news`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `language`
-- 
DROP TABLE IF EXISTS `language`;
CREATE TABLE IF NOT EXISTS `language` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `descrizione` varchar(250) default NULL,
  `label` varchar(100) default NULL,
  `lang_active` int(1) unsigned NOT NULL default '0',
  `subdomain_active` int(1) unsigned NOT NULL default '0',
  `url_subdomain` varchar(250) default NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `language_disponibili`
-- 
DROP TABLE IF EXISTS `language_disponibili`;
CREATE TABLE IF NOT EXISTS `language_disponibili` (
  `keyword` VARCHAR(2) NOT NULL,
  `description` VARCHAR(45) NOT NULL,
  PRIMARY KEY (`keyword`)
) ENGINE = InnoDB  DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `multi_languages`
-- 
DROP TABLE IF EXISTS `multi_languages`;
CREATE TABLE IF NOT EXISTS `multi_languages` (
  `id` int(20) unsigned NOT NULL auto_increment,
  `keyword` varchar(150) NOT NULL,
  `lang_code` varchar(10) NOT NULL,
  `value` text,
  PRIMARY KEY  (`id`),
  UNIQUE KEY `Index_ML` (`keyword`,`lang_code`),
  INDEX `Index_kw`(`keyword`),
  INDEX `Index_lc`(`lang_code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `logs`
-- 
DROP TABLE IF EXISTS `logs`;
CREATE TABLE IF NOT EXISTS `logs` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `msg` TEXT default NULL,
  `usr` varchar(50) NOT NULL,
  `type` varchar(15) NOT NULL,
  `date_event` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`id`),
  KEY `usr` (`usr`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `news`
-- 
DROP TABLE IF EXISTS `news`;
CREATE TABLE IF NOT EXISTS `news` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `titolo` varchar(250) NOT NULL,
  `abstract` text,
  `abstract_2` text,
  `abstract_3` text,
  `testo` text,
  `keyword` varchar(100) default NULL,
  `data_inserimento` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `data_pubblicazione` timestamp NOT NULL default '0000-00-00 00:00:00',
  `data_cancellazione` timestamp NOT NULL default '0000-00-00 00:00:00',
  `stato_news` int(2) unsigned NOT NULL,
  `meta_description` TEXT default NULL,
  `meta_keyword` TEXT default NULL,
  `page_title` TEXT default NULL,
  PRIMARY KEY  (`id`),
  KEY `Index_2` (`titolo`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `newsletter`
-- 
DROP TABLE IF EXISTS `newsletter`;
CREATE TABLE IF NOT EXISTS `newsletter` (
  `id_newsletter` int(10) unsigned NOT NULL auto_increment,
  `descrizione` varchar(100) NOT NULL,
  `stato` int(10) unsigned NOT NULL default '0',
  `template` varchar(100) NOT NULL,
  `id_voucher_campaign` int(10) unsigned default NULL,
  PRIMARY KEY  (`id_newsletter`),
  KEY `Index_stato` (`stato`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `newsletter_x_utente`
-- 
DROP TABLE IF EXISTS `newsletter_x_utente`;
CREATE TABLE `newsletter_x_utente` (
  `id_newsletter` int(10) unsigned NOT NULL,
  `id_utente` int(10) unsigned NOT NULL,
  KEY `Index_1` (`id_utente`),
  KEY `Index_2` (`id_newsletter`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `news_x_utente`
-- 
DROP TABLE IF EXISTS `news_x_utente`;
CREATE TABLE IF NOT EXISTS `news_x_utente` (
  `id_news` int(10) unsigned NOT NULL,
  `id_utente` int(10) unsigned NOT NULL,
  KEY `Index_1` (`id_news`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
--
-- Struttura della tabella `target`
-- 
DROP TABLE IF EXISTS `target`;
CREATE TABLE IF NOT EXISTS `target` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `descrizione` varchar(250) default NULL,
  `type` int(1) unsigned NOT NULL,
  `locked` smallint(1) unsigned NOT NULL default '0',
  `automatic` smallint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (`id`),
  KEY `descrizione` (`descrizione`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `target_type`
-- 
DROP TABLE IF EXISTS `target_type`;
CREATE TABLE IF NOT EXISTS `target_type` (
  `id` int(10) unsigned NOT NULL,
  `descrizione` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `target_x_categoria`
-- 
DROP TABLE IF EXISTS `target_x_categoria`;
CREATE TABLE IF NOT EXISTS `target_x_categoria` (
  `id_target` int(10) unsigned NOT NULL,
  `id_categoria` int(10) unsigned NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `target_x_news`
-- 
DROP TABLE IF EXISTS `target_x_news`;
CREATE TABLE IF NOT EXISTS `target_x_news` (
  `id_target` int(10) unsigned NOT NULL,
  `id_news` int(10) unsigned NOT NULL,
  KEY `Index_2` (`id_news`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `target_x_utente`
-- 
DROP TABLE IF EXISTS `target_x_utente`;
CREATE TABLE IF NOT EXISTS `target_x_utente` (
  `id_target` int(10) unsigned NOT NULL,
  `id_utente` int(10) unsigned NOT NULL,
  KEY `Index_1` USING BTREE (`id_utente`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `template_disponibili`
-- 
DROP TABLE IF EXISTS `template_disponibili`;
CREATE TABLE IF NOT EXISTS `template_disponibili` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `dir_template` varchar(250) NOT NULL,
  `template_css` varchar(100) default NULL,
  `descrizione` varchar(250) default NULL,
  `base_template` int(2) unsigned NOT NULL,
  `order_by` int(2) unsigned NOT NULL,
  `elem_x_page` int(3) unsigned NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `page_x_template`
-- 
DROP TABLE IF EXISTS `page_x_template`;
CREATE TABLE IF NOT EXISTS `page_x_template` (
  `id` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
  `id_template` INTEGER UNSIGNED NOT NULL,
  `file_name` VARCHAR(100) NOT NULL,
  `page_num` INTEGER NOT NULL,
  PRIMARY KEY (`id`),
  INDEX `Index_2`(`id_template`),
  INDEX `Index_3`(`page_num`)
)ENGINE = InnoDB  DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `template_x_categoria`
-- 
DROP TABLE IF EXISTS `template_x_categoria`;
CREATE TABLE IF NOT EXISTS `template_x_categoria` (
  `id_categoria` INTEGER UNSIGNED NOT NULL,
  `id_template` INTEGER UNSIGNED NOT NULL,
  `lang_code` varchar(10) NOT NULL,
  INDEX `Index_2`(`id_categoria`),
  INDEX `Index_3`(`id_template`),
  INDEX `Index_4`(`lang_code`)
)ENGINE = InnoDB  DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `uploaded_files`
-- 
DROP TABLE IF EXISTS `uploaded_files`;
CREATE TABLE IF NOT EXISTS `uploaded_files` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `filename` varchar(250) NOT NULL,
  `content_type` varchar(50) NOT NULL,
  `path` varchar(250) NOT NULL,
  `file_dida` text,
  `file_label` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `downloaded_files`
-- 
DROP TABLE IF EXISTS `downloaded_files`;
CREATE TABLE IF NOT EXISTS `downloaded_files` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `id_file` INT(11) UNSIGNED NOT NULL ,
  `id_user` INT(11) UNSIGNED default NULL ,
  `user_host` varchar(100) default NULL,
  `user_info` varchar(250) default NULL,
  `filename` varchar(250) NOT NULL,
  `content_type` varchar(50) NOT NULL,
  `path` varchar(250) NOT NULL,
  `download_date` TIMESTAMP NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `user_photos`
-- 
DROP TABLE IF EXISTS `user_files`;
CREATE TABLE IF NOT EXISTS `user_files` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `id_user` INTEGER UNSIGNED NOT NULL,
  `filename` varchar(250) NOT NULL,
  `content_type` varchar(50) NOT NULL,
  `path` varchar(250) NOT NULL,
  `file_dida` text,
  `file_label` varchar(100) NOT NULL,
  `dta_ins` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `utenti`
-- 
DROP TABLE IF EXISTS `utenti`;
CREATE TABLE IF NOT EXISTS `utenti` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `username` varchar(100) NOT NULL,
  `pwd` varchar(50) NOT NULL,
  `email` varchar(100) NOT NULL,
  `ruolo` int(10) unsigned NOT NULL,
  `privacy` char(1) NOT NULL,
  `newsletter` char(1) NOT NULL,
  `utenteAttivo` varchar(2) NOT NULL,
  `sconto` DECIMAL(10,2) NOT NULL default '0',
  `adminComments` text,
  `insertDate` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `modifyDate` timestamp NOT NULL default '0000-00-00 00:00:00',
  `public` INTEGER(1) UNSIGNED NOT NULL DEFAULT 0,
  `user_group` int(11) unsigned default NULL,  
  `automatic_user` smallint(1) unsigned NOT NULL DEFAULT '0',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `utenti_images`
-- 
DROP TABLE IF EXISTS `utenti_images`;
CREATE TABLE IF NOT EXISTS `utenti_images` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `id_utente` int(10) unsigned NOT NULL,
  `filename` varchar(250) NOT NULL,
  `content_type` varchar(50) NOT NULL,
  `file_size` bigint(11) unsigned NOT NULL,
  `file_data` blob,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `ruoli_utente`
-- 
DROP TABLE IF EXISTS `ruoli_utente`;
CREATE TABLE IF NOT EXISTS `ruoli_utente` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `descrizione` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `user_preference`
-- 
DROP TABLE IF EXISTS `user_preference`;
CREATE TABLE `user_preference` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `id_user` int(10) unsigned NOT NULL,
  `id_friend` int(10) unsigned NOT NULL,
  `id_usr_comment` int(10) unsigned default NULL,
  `comment_type` int(2) unsigned default NULL,
  `type` int(1) NOT NULL,
  `value` text,
  `dta_insert` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`id`),
  KEY `Index_2` (`id_user`),
  KEY `Index_3` (`id_friend`),
  KEY `Index_4` (`type`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `friend_x_user`
-- 
DROP TABLE IF EXISTS `friend_x_user`;
CREATE TABLE `friend_x_user` (
  `id_friend` INTEGER UNSIGNED NOT NULL,
  `id_user` INTEGER UNSIGNED NOT NULL,
  `active` int(1) UNSIGNED NOT NULL DEFAULT 0,
  PRIMARY KEY (`id_friend`, `id_user`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `user_fields`
-- 
DROP TABLE IF EXISTS `user_fields`;
CREATE TABLE `user_fields` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  `id_group` int(11) unsigned DEFAULT NULL,
  `type` int(11) unsigned NOT NULL,
  `type_content` int(11) unsigned NOT NULL,
  `values` text,
  `order` int(2) unsigned NOT NULL DEFAULT 0,
  `required` int(1) UNSIGNED NOT NULL DEFAULT 0,
  `enabled` int(1) UNSIGNED NOT NULL DEFAULT 0,
  `max_lenght` int(3) UNSIGNED DEFAULT NULL,
  `use_for` int(3) UNSIGNED DEFAULT NULL,
  PRIMARY KEY  (`id`),
  KEY `Index_3` (`id_group`),
  KEY `Index_4` (`type`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `user_fields_group`
-- 
DROP TABLE IF EXISTS `user_fields_group`;
CREATE TABLE `user_fields_group` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  `order` int(2) unsigned NOT NULL DEFAULT 0,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `user_fields_type`
-- 
DROP TABLE IF EXISTS `user_fields_type`;
CREATE TABLE `user_fields_type` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `user_fields_type_content`
-- 
DROP TABLE IF EXISTS `user_fields_type_content`;
CREATE TABLE `user_fields_type_content` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  `description` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `user_fields_match`
-- 
DROP TABLE IF EXISTS `user_fields_match`;
CREATE TABLE `user_fields_match` (
  `id_field` INTEGER UNSIGNED NOT NULL,
  `id_user` INTEGER UNSIGNED NOT NULL,
  `value` text,
  PRIMARY KEY (`id_field`, `id_user`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `countries`
-- 
DROP TABLE IF EXISTS `countries`;
CREATE TABLE IF NOT EXISTS `countries` (
`id` int(10) unsigned NOT NULL auto_increment,
  `country_code` VARCHAR(2) NOT NULL,
  `country_description` VARCHAR(100) NOT NULL,
  `state_region_code` VARCHAR(10) DEFAULT NULL,
  `state_region_description` VARCHAR(100) DEFAULT NULL,
  `active` SMALLINT(1) UNSIGNED NOT NULL default '0',
  `use_for` int(3) UNSIGNED DEFAULT NULL,  
  PRIMARY KEY (`id`),
  INDEX `Index_CC`(`country_code`),
  INDEX `Index_SRC`(`state_region_code`)
) ENGINE = InnoDB  DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `googlemap_localization`
-- 
DROP TABLE IF EXISTS `googlemap_localization`;
CREATE TABLE IF NOT EXISTS `googlemap_localization` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `id_element` int(10) NOT NULL,
  `type` SMALLINT(1) UNSIGNED NOT NULL default '1',
  `latitude` decimal(10,6) DEFAULT NULL,
  `longitude` decimal(10,6) DEFAULT NULL, 
  `txtinfo` text,
  PRIMARY KEY (`id`),
  UNIQUE KEY `Index_gl` (`id_element`,`type`,`latitude`,`longitude`)
) ENGINE = InnoDB  DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `ads`
-- 
DROP TABLE IF EXISTS `ads`;
CREATE TABLE IF NOT EXISTS `ads` (
  `id_ads` int(10) unsigned NOT NULL auto_increment,
  `id_element` int(10) unsigned NOT NULL,
  `id_utente` int(10) NOT NULL,
  `phone` varchar(100) DEFAULT NULL,
  `ads_type` int(1) unsigned NOT NULL default '0',
  `price` decimal(10,2) DEFAULT NULL,
  `dta_inserimento` timestamp NOT NULL default CURRENT_TIMESTAMP,
  PRIMARY KEY  (`id_ads`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `content_fields`
-- 
DROP TABLE IF EXISTS `content_fields`;
CREATE TABLE `content_fields` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  `id_group` int(11) unsigned DEFAULT NULL,
  `type` int(11) unsigned NOT NULL,
  `type_content` int(11) unsigned NOT NULL,
  `order` int(3) unsigned NOT NULL DEFAULT 0,
  `required` int(1) UNSIGNED NOT NULL DEFAULT 0,
  `enabled` int(1) UNSIGNED NOT NULL DEFAULT 0,
  `max_lenght` int(3) UNSIGNED DEFAULT NULL,
  `editable` int(1) UNSIGNED NOT NULL DEFAULT 0,
  PRIMARY KEY  (`id`),
  KEY `Index_3` (`id_group`),
  KEY `Index_4` (`type`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `content_fields_values`
-- 
DROP TABLE IF EXISTS `content_fields_values`;
CREATE TABLE `content_fields_values` (
  `id_field` int(11) unsigned NOT NULL,
  `value` varchar(250) NOT NULL,
  `order` int(3) unsigned NOT NULL DEFAULT 0,
  UNIQUE KEY `Index_PFV` (`id_field`,`value`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `content_fields_group`
-- 
DROP TABLE IF EXISTS `content_fields_group`;
CREATE TABLE `content_fields_group` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  `order` int(2) unsigned NOT NULL DEFAULT 0,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `content_fields_type`
-- 
DROP TABLE IF EXISTS `content_fields_type`;
CREATE TABLE `content_fields_type` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `content_fields_type_content`
-- 
DROP TABLE IF EXISTS `content_fields_type_content`;
CREATE TABLE `content_fields_type_content` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `content_fields_match`
-- 
DROP TABLE IF EXISTS `content_fields_match`;
CREATE TABLE `content_fields_match` (
  `id_field` INTEGER UNSIGNED NOT NULL,
  `id_news` INTEGER UNSIGNED NOT NULL,
  `value` varchar(250) NOT NULL,
  PRIMARY KEY (`id_field`, `id_news`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- ----------------------------------------------------------------------- -- ---------------------------------------- -- ----------------------------------------------------------------------- --
-- ----------------------------------------------------------------------- -- TABELLE ECONEME-SYS -- ----------------------------------------------------------------------- --
-- ----------------------------------------------------------------------- -- ---------------------------------------- -- ----------------------------------------------------------------------- --


-- --------------------------------------------------------
-- 
-- Struttura della tabella `ads_promotion`
-- 
DROP TABLE IF EXISTS `ads_promotion`;
CREATE TABLE IF NOT EXISTS `ads_promotion` (
  `id_ads` int(10) unsigned NOT NULL,
  `id_element` int(10) unsigned NOT NULL,
  `cod_element` VARCHAR(100) NOT NULL,
  `active` SMALLINT(1) UNSIGNED NOT NULL default '0',
  `dta_inserimento` timestamp NOT NULL default CURRENT_TIMESTAMP,
  PRIMARY KEY  (`id_ads`,`id_element`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `ordini`
-- 
DROP TABLE IF EXISTS `ordini`;
CREATE TABLE IF NOT EXISTS `ordini` (
  `id_ordine` int(10) unsigned NOT NULL auto_increment,
  `id_utente` int(10) unsigned NOT NULL,
  `dta_inserimento` timestamp NOT NULL default CURRENT_TIMESTAMP,
  `stato_ordine` varchar(100) NOT NULL,
  `totale_imponibile` DECIMAL(10,2) NOT NULL,
  `totale_tasse` DECIMAL(10,2) NOT NULL,
  `totale` decimal(10,2) NOT NULL,
  `tipo_pagam` varchar(100),
  `payment_commission` decimal(10,2) NOT NULL default '0.00',
  `pagam_effettuato` int(10) unsigned NOT NULL,
  `order_guid` varchar(250) NOT NULL,
  `user_notified_x_download` INT(1) UNSIGNED NOT NULL DEFAULT '0',
  `notes` text,
  `no_registration` smallint(1) unsigned NOT NULL DEFAULT '0',
  `id_ads` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY  (`id_ordine`),
  KEY `Index_user` (`id_utente`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `prodotti`
-- 
DROP TABLE IF EXISTS `prodotti`;
CREATE TABLE IF NOT EXISTS `prodotti` (
  `id_prodotto` int(10) unsigned NOT NULL auto_increment,
  `nome_prod` varchar(250) NOT NULL,
  `sommario_prod` text,
  `desc_prod` text,
  `prezzo` decimal(10,2) NOT NULL,
  `qta_disp` varchar(100) NOT NULL,
  `attivo` int(10) unsigned NOT NULL,
  `sconto` DECIMAL(10,2) NOT NULL default '0',
  `codice_prod` varchar(100) NOT NULL,
  `id_tassa_applicata` int(10) unsigned default NULL,
  `prod_type` smallint(1) unsigned NOT NULL,
  `max_download` int(11) NOT NULL default '-1',
  `max_download_time` int(11) NOT NULL default '-1',
  `taxs_group` INT( 10 ) UNSIGNED DEFAULT NULL ,
  `meta_description` TEXT default NULL,
  `meta_keyword` TEXT default NULL,
  `page_title` TEXT default NULL,
  `edit_buy_qta` smallint(1) unsigned NOT NULL default '0',  
  PRIMARY KEY  (`id_prodotto`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `attach_x_prodotti`
-- 
DROP TABLE IF EXISTS `attach_x_prodotti`;
CREATE TABLE IF NOT EXISTS `attach_x_prodotti` (
  `id_prodotto` int(10) unsigned NOT NULL,
  `id_attach` int(10) unsigned NOT NULL auto_increment,
  `filename` varchar(100) NOT NULL,
  `content_type` varchar(20) NOT NULL,
  `path` varchar(100) NOT NULL,
  `file_dida` text,
  `file_label` varchar(2) NOT NULL,
  PRIMARY KEY  (`id_attach`),
  KEY `Index_2` (`id_prodotto`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `relation_x_prodotto`
-- 
DROP TABLE IF EXISTS `relation_x_prodotto`;
CREATE TABLE IF NOT EXISTS `relation_x_prodotto` (
  `id_prod` int(10) unsigned NOT NULL,
  `id_prod_rel` int(10) unsigned NOT NULL,
  UNIQUE KEY `Index_Rp` (`id_prod`,`id_prod_rel`),
  INDEX `Index_RpP`(`id_prod`),
  INDEX `Index_RpPr`(`id_prod_rel`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `prodotto_main_field_translation`
-- 
DROP TABLE IF EXISTS `prodotto_main_field_translation`;
CREATE TABLE IF NOT EXISTS `prodotto_main_field_translation` (
  `id_prod` int(10) unsigned NOT NULL,
  `main_field` int(3) unsigned NOT NULL,
  `lang_code` varchar(2) NOT NULL,
  `value` text,
  UNIQUE KEY `Index_Pmft` (`id_prod`,`main_field`,`lang_code`),
  INDEX `Index_Pmfti`(`id_prod`),
  INDEX `Index_Pmftm`(`main_field`),
  INDEX `Index_Pmftl`(`lang_code`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `carrello`
-- 
DROP TABLE IF EXISTS `carrello`;
CREATE TABLE IF NOT EXISTS `carrello` (
  `id_carrello` int(10) unsigned NOT NULL auto_increment,
  `id_utente` int(11) NOT NULL,
  `dta_creazione` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`id_carrello`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `downloadable_products`
-- 
DROP TABLE IF EXISTS `downloadable_products`;
CREATE TABLE IF NOT EXISTS `downloadable_products` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `id_product` int(10) unsigned NOT NULL,
  `filename` varchar(250) NOT NULL,
  `path` varchar(250) NOT NULL,
  `content_type` varchar(50) NOT NULL,
  `file_size` int(10) unsigned NOT NULL,
  `insert_date` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`id`),
  KEY `Index_2` (`id_product`),
  KEY `Index_3` (`filename`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- --------------------------------------------------------
-- 
-- Struttura della tabella `down_prod_x_order`
-- 
DROP TABLE IF EXISTS `down_prod_x_order`;
CREATE TABLE `down_prod_x_order` (
`id` INT(11) UNSIGNED NOT NULL AUTO_INCREMENT,
`id_order` INT(11) UNSIGNED NOT NULL ,
`id_prod` INT(11) UNSIGNED NOT NULL ,
`id_down_prod` INT(11) UNSIGNED NOT NULL ,
`id_user` INT(11) UNSIGNED NOT NULL ,
`active` SMALLINT(1) UNSIGNED NOT NULL default '0',
`max_num_download` INT(3) NOT NULL  default '-1',
`insert_date` TIMESTAMP NOT NULL ,
`expire_date` TIMESTAMP NULL ,
`download_counter` INT(3) UNSIGNED NOT NULL  default '0',
`download_date` TIMESTAMP NULL,
 PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- --------------------------------------------------------
-- 
-- Struttura della tabella `prodotti_x_carrello`
-- 
DROP TABLE IF EXISTS `prodotti_x_carrello`;
CREATE TABLE IF NOT EXISTS `prodotti_x_carrello` (
  `id_carrello` int(10) unsigned NOT NULL,
  `id_prodotto` int(10) unsigned NOT NULL,
  `counter_prod` int(10) unsigned NOT NULL,
  `qta_prod` int(10) unsigned NOT NULL,
  `prod_type` smallint(1) unsigned NOT NULL,
  PRIMARY KEY  (`id_carrello`,`id_prodotto`,`counter_prod`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `prodotti_x_ordine`
-- 
DROP TABLE IF EXISTS `prodotti_x_ordine`;
CREATE TABLE IF NOT EXISTS `prodotti_x_ordine` (
  `id_ordine` int(10) unsigned NOT NULL,
  `id_prodotto` int(10) unsigned NOT NULL,
  `counter_prod` int(10) unsigned NOT NULL,
  `nome_prodotto` varchar(100) NOT NULL,
  `qta` int(10) unsigned NOT NULL,
  `totale` decimal(10,2) NOT NULL,
  `tax` decimal(10,2) NOT NULL default '0.00',  
  `desc_tax` varchar(100) DEFAULT NULL,
  `prod_type` smallint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (`id_ordine`,`id_prodotto`,`counter_prod`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `spese_x_ordine`
-- 
DROP TABLE IF EXISTS `spese_x_ordine`;
CREATE TABLE `spese_x_ordine` (
  `id_ordine` INTEGER UNSIGNED NOT NULL,
  `id_spesa` INTEGER UNSIGNED NOT NULL,
  `imponibile` DECIMAL(10,2) NOT NULL,
  `tasse` DECIMAL(10,2) NOT NULL,
  `totale` DECIMAL(10,2) NOT NULL, 
  `desc_spesa` varchar(100) DEFAULT NULL,
  INDEX `Index_1`(`id_ordine`),
  INDEX `Index_2`(`id_spesa`)
)ENGINE = InnoDB  DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `shipping_address`
-- 
DROP TABLE IF EXISTS `shipping_address`;
CREATE TABLE IF NOT EXISTS `shipping_address` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `id_user` int(10) unsigned NOT NULL,
  `name` varchar(100) default NULL,
  `surname` varchar(100) default NULL,
  `cfiscvat` varchar(16) default NULL,
  `address` varchar(250) default NULL,
  `city` varchar(100) default NULL,
  `zipCode` varchar(20) default NULL,
  `country` varchar(100) default NULL,
  `state_region` varchar(100) default NULL,
  `is_company_client` SMALLINT(1) UNSIGNED NOT NULL default '0',  
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `order_shipping_address`
-- 
DROP TABLE IF EXISTS `order_shipping_address`;
CREATE TABLE IF NOT EXISTS `order_shipping_address` (
  `id_order` int(10) unsigned NOT NULL,
  `id_shipping` int(10) unsigned NOT NULL,
  `address` varchar(250) default NULL,
  `city` varchar(100) default NULL,
  `zipCode` varchar(20) default NULL,
  `country` varchar(100) default NULL,
  `state_region` varchar(100) default NULL,
  `is_company_client` SMALLINT(1) UNSIGNED NOT NULL default '0'
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `bills_address`
-- 
DROP TABLE IF EXISTS `bills_address`;
CREATE TABLE IF NOT EXISTS `bills_address` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `id_user` int(10) unsigned NOT NULL,
  `name` varchar(100) default NULL,
  `surname` varchar(100) default NULL,
  `cfiscvat` varchar(16) default NULL,
  `address` varchar(250) default NULL,
  `city` varchar(100) default NULL,
  `zipCode` varchar(20) default NULL,
  `country` varchar(100) default NULL,
  `state_region` varchar(100) default NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `order_bills_address`
-- 
DROP TABLE IF EXISTS `order_bills_address`;
CREATE TABLE IF NOT EXISTS `order_bills_address` (
  `id_order` int(10) unsigned NOT NULL,
  `id_bills` int(10) unsigned NOT NULL,
  `address` varchar(250) default NULL,
  `city` varchar(100) default NULL,
  `zipCode` varchar(20) default NULL,
  `country` varchar(100) default NULL,
  `state_region` varchar(100) default NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `target_x_prodotto`
-- 
DROP TABLE IF EXISTS `target_x_prodotto`;
CREATE TABLE IF NOT EXISTS `target_x_prodotto` (
  `id_target` int(10) unsigned NOT NULL,
  `id_prodotto` int(10) unsigned NOT NULL,
  KEY `Index_1` (`id_prodotto`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `spese_accessorie`
-- 
DROP TABLE IF EXISTS `spese_accessorie`;
CREATE TABLE IF NOT EXISTS `spese_accessorie` (
  `id` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
  `descrizione_spesa` VARCHAR(100) NOT NULL,
 `valore` DECIMAL(10,2) NOT NULL,
  `tipologia_valore` SMALLINT(1) UNSIGNED NOT NULL,
  `id_tassa_applicata` INTEGER(10) UNSIGNED default NULL,
  `applica_frontend`SMALLINT(1) UNSIGNED,
  `applica_backend` SMALLINT(1) UNSIGNED,
  `autoactive` SMALLINT(1) UNSIGNED NOT NULL default '0',
  `multiply` SMALLINT(1) UNSIGNED NOT NULL default '0',
  `required` SMALLINT(1) UNSIGNED NOT NULL default '0',
  `group` VARCHAR(50) NOT NULL,
  `taxs_group` INT( 10 ) UNSIGNED DEFAULT NULL ,  
  `type_view` SMALLINT(1) UNSIGNED NOT NULL default '0',
  PRIMARY KEY (`id`),
  INDEX `Index_2`(`valore`),
  INDEX `Index_3`(`id_tassa_applicata`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `spese_accessorie_config`
-- 
DROP TABLE IF EXISTS `spese_accessorie_config`;
CREATE TABLE IF NOT EXISTS `spese_accessorie_config` (
  `id` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
  `id_spesa` int(10) unsigned NOT NULL,
  `id_prod_field` int(10) unsigned DEFAULT NULL,
 `rate_from` DECIMAL(10,2) NOT NULL,
 `rate_to` DECIMAL(10,2) NOT NULL,
 `operation` SMALLINT( 1 ) UNSIGNED NOT NULL DEFAULT '0' COMMENT 'tipo di operatione da eseguire: 0 nulla, 1 somma, 2 sottrazione;',
 `valore` DECIMAL(10,2) NOT NULL,  
  PRIMARY KEY (`id`),
  UNIQUE KEY `Index_U` (`id_spesa`,`id_prod_field`,`rate_from`,`rate_to`),
  INDEX `Index_From`(`rate_from`),
  INDEX `Index_To`(`rate_to`),
  INDEX `Index_Val`(`valore`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;



-- --------------------------------------------------------
-- 
-- Struttura della tabella `tasse`
-- 
DROP TABLE IF EXISTS `tasse`;
CREATE TABLE `tasse` (
  `id` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
  `descrizione_tassa` VARCHAR(100) NOT NULL,
  `valore` DECIMAL(10,2) NOT NULL,
  `tipologia_valore` SMALLINT(1) UNSIGNED NOT NULL,
  PRIMARY KEY (`id`),
  INDEX `Index_2`(`valore`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella tax_group`
-- 
DROP TABLE IF EXISTS `tax_group`;
CREATE TABLE IF NOT EXISTS `tax_group` (  
	  `id` int(10) unsigned NOT NULL auto_increment,
  `description` VARCHAR(100) NOT NULL,
  PRIMARY KEY (`id`),
  INDEX `Index_TG_dc`(`description`)
) ENGINE = InnoDB  DEFAULT CHARSET=utf8;	

-- --------------------------------------------------------
-- 
-- Struttura della tabella tax_group_value`
-- 
DROP TABLE IF EXISTS `tax_group_value`;
CREATE TABLE IF NOT EXISTS `tax_group_value` (  
  `id_group` int(10) unsigned NOT NULL,
  `country_code` VARCHAR(2) NOT NULL,
  `state_region_code` VARCHAR(10) DEFAULT NULL,
  `id_tassa_applicata` int(10) unsigned default NULL,
  `exclude_calculation` SMALLINT(1) UNSIGNED NOT NULL default '0',
  INDEX `Index_TGV_ig`(`id_group`),		  
  INDEX `Index_TGV_cc`(`country_code`),		  
  INDEX `Index_TGV_src`(`state_region_code`)
) ENGINE = InnoDB  DEFAULT CHARSET=utf8;	
		

-- --------------------------------------------------------
-- 
-- Struttura della tabella `payment` e tabelle correlate
-- 
DROP TABLE IF EXISTS `payment_type`;
CREATE TABLE `payment_type` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `keyword_multilingua` varchar(250) default NULL,
  `descrizione` varchar(250) default NULL,
  `dati_pagamento` varchar(250) NOT NULL,
  `commission` decimal(10,2) NOT NULL default '0.00',
  `commission_type` SMALLINT(1) UNSIGNED NOT NULL  default '1',
  `url` smallint(1) unsigned NOT NULL default '0',
  `id_modulo` int(10) default NULL,
  `activate` smallint(1) unsigned NOT NULL default '0',
  `payment_type` smallint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

DROP TABLE IF EXISTS `payment_field`;
CREATE TABLE `payment_field` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `id_payment` int(10) unsigned NOT NULL,
  `id_modulo` int(10) unsigned default NULL,
  `name` varchar(50) NOT NULL,
  `value` varchar(250) default NULL,
  `match_field` varchar(50) default NULL,
  PRIMARY KEY  USING BTREE (`id`),
  UNIQUE KEY `Index_UX` (`id_payment`,`id_modulo`,`name`,`match_field`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

DROP TABLE IF EXISTS `payment_modulo`;
CREATE TABLE `payment_modulo` (
  `id` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
  `name` VARCHAR(45) NOT NULL,
  `directory` VARCHAR(100) NOT NULL,
  `logo` TEXT,
  `insert_page` VARCHAR(100) NOT NULL,
  `checkout_page` VARCHAR(100) NOT NULL,
  `checkin_page` VARCHAR(100) NOT NULL,
  `checkin_fault_page` VARCHAR(100) NOT NULL,
  `id_ordine_field` VARCHAR(100) NOT NULL,
  `ip_provider` VARCHAR(150) NOT NULL,
  PRIMARY KEY (`id`),
  INDEX `Index_2`(`name`)
)ENGINE = InnoDB  DEFAULT CHARSET=utf8;

DROP TABLE IF EXISTS `payment_fixed_app_field`;
CREATE TABLE `payment_fixed_app_field` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `keyword` varchar(50) NOT NULL,
  `value` varchar(100) default NULL,
  `used` smallint(1) unsigned NOT NULL default '1',
  PRIMARY KEY  (`id`),
  UNIQUE KEY `Index_2` (`keyword`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

DROP TABLE IF EXISTS `paypal_field`;
CREATE TABLE `paypal_field` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `keyword` varchar(50) NOT NULL,
  `value` varchar(100) default NULL,
  `match_field` varchar(50) default NULL,
  PRIMARY KEY  (`id`),
  UNIQUE KEY `Index_2` (`keyword`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

DROP TABLE IF EXISTS `sella_field`;
CREATE TABLE `sella_field` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `keyword` varchar(50) NOT NULL,
  `value` varchar(100) default NULL,
  `match_field` varchar(50) default NULL,
  PRIMARY KEY  (`id`),
  UNIQUE KEY `Index_2` (`keyword`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

DROP TABLE IF EXISTS `payment_transactions`;
CREATE TABLE `payment_transactions` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `id_ordine` int(11) unsigned NOT NULL,
  `id_modulo` INTEGER UNSIGNED NOT NULL,
  `id_transaction` varchar(100) NOT NULL,
  `status` varchar(50) default NULL,
  `notified` smallint(1) unsigned NOT NULL default '0',
  `insert_date` datetime NOT NULL,
  PRIMARY KEY  (`id`),
  INDEX `Index_2` (`id_ordine`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `currency`
-- 
DROP TABLE IF EXISTS `currency`;
CREATE TABLE `currency` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `currency` varchar(5) NOT NULL,
  `rate` decimal(10,4) NOT NULL,
  `dta_riferimento` date NOT NULL,
  `dta_inserimento` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `active` int(1) unsigned NOT NULL default '0',
  `is_default` int(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (`id`),
  KEY `currency` (`currency`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `user_group`
-- 
DROP TABLE IF EXISTS `user_group`;
CREATE TABLE IF NOT EXISTS `user_group` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `short_desc` varchar(100) NOT NULL,
  `long_desc` text,
  `default_group` int(1) unsigned NOT NULL default '0',
  `taxs_group` INT( 10 ) UNSIGNED DEFAULT NULL ,  
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `margin_discount`
-- 
DROP TABLE IF EXISTS `margin_discount`;
CREATE TABLE IF NOT EXISTS `margin_discount` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `margin` decimal(10,2) unsigned NOT NULL,
  `discount` decimal(10,2) unsigned NOT NULL,
  `apply_prod_discount` int(1) unsigned NOT NULL default '0',
  `apply_user_discount` int(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `usr_group_x_margin_disc`
-- 
DROP TABLE IF EXISTS `usr_group_x_margin_disc`;
CREATE TABLE IF NOT EXISTS `usr_group_x_margin_disc` (
  `id_marg_disc` int(11) unsigned NOT NULL,
  `id_user_group` int(11) unsigned NOT NULL,
  KEY `Index_1` (`id_user_group`),
  UNIQUE KEY `Index_2` (`id_user_group`),  
  UNIQUE KEY `Index_3` (`id_user_group`,`id_marg_disc`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `product_fields`
-- 
DROP TABLE IF EXISTS `product_fields`;
CREATE TABLE `product_fields` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  `id_group` int(11) unsigned DEFAULT NULL,
  `type` int(11) unsigned NOT NULL,
  `type_content` int(11) unsigned NOT NULL,
  `order` int(3) unsigned NOT NULL DEFAULT 0,
  `required` int(1) UNSIGNED NOT NULL DEFAULT 0,
  `enabled` int(1) UNSIGNED NOT NULL DEFAULT 0,
  `max_lenght` int(3) UNSIGNED DEFAULT NULL,
  `editable` int(1) UNSIGNED NOT NULL DEFAULT 0,
  PRIMARY KEY  (`id`),
  KEY `Index_3` (`id_group`),
  KEY `Index_4` (`type`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `product_fields_values`
-- 
DROP TABLE IF EXISTS `product_fields_values`;
CREATE TABLE `product_fields_values` (
  `id_field` int(11) unsigned NOT NULL,
  `value` varchar(250) NOT NULL,
  `order` int(3) unsigned NOT NULL DEFAULT 0,
  UNIQUE KEY `Index_PFV` (`id_field`,`value`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `product_fields_group`
-- 
DROP TABLE IF EXISTS `product_fields_group`;
CREATE TABLE `product_fields_group` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  `order` int(2) unsigned NOT NULL DEFAULT 0,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `product_fields_type`
-- 
DROP TABLE IF EXISTS `product_fields_type`;
CREATE TABLE `product_fields_type` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `product_fields_type_content`
-- 
DROP TABLE IF EXISTS `product_fields_type_content`;
CREATE TABLE `product_fields_type_content` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `description` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `product_fields_match`
-- 
DROP TABLE IF EXISTS `product_fields_match`;
CREATE TABLE `product_fields_match` (
  `id_field` INTEGER UNSIGNED NOT NULL,
  `id_prod` INTEGER UNSIGNED NOT NULL,
  `value` varchar(250) NOT NULL,
  PRIMARY KEY (`id_field`, `id_prod`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `product_fields_value_match`
-- 
DROP TABLE IF EXISTS `product_fields_value_match`;
CREATE TABLE `product_fields_value_match` (
  `id_field` INTEGER UNSIGNED NOT NULL,
  `id_prod` INTEGER UNSIGNED NOT NULL,
  `qta_prod` int(10) NOT NULL,
  `value` varchar(250) NOT NULL,
  PRIMARY KEY (`id_field`, `id_prod`, `value`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `product_fields_rel_value_match`
-- 
DROP TABLE IF EXISTS `product_fields_rel_value_match`;
CREATE TABLE IF NOT EXISTS `product_fields_rel_value_match` (
  `id_prod` int(10) unsigned NOT NULL,
  `id_field` int(10) unsigned NOT NULL,
  `field_val` varchar(250) NOT NULL,
  `id_field_rel` int(10) unsigned NOT NULL,
  `field_rel_val` varchar(250) NOT NULL,
  `qta_rel` int(10) NOT NULL,
  KEY `id_prod` (`id_prod`),
  KEY `id_field` (`id_field`),
  KEY `field_val` (`field_val`),
  KEY `id_field_rel` (`id_field_rel`),
  KEY `field_rel_val` (`field_rel_val`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `product_fields_x_order`
-- 
DROP TABLE IF EXISTS `product_fields_x_order`;
CREATE TABLE IF NOT EXISTS `product_fields_x_order` (
  `counter` INTEGER UNSIGNED NOT NULL,
  `id_order` int(10) unsigned NOT NULL,
  `id_prod` int(10) unsigned NOT NULL,
  `id_field` int(10) unsigned NOT NULL,
  `qta_prod` int(10) unsigned NOT NULL,
  `value` varchar(250) NOT NULL,
  PRIMARY KEY  (`counter`,`id_order`,`id_prod`,`id_field`,`value`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `product_fields_x_card`
-- 
DROP TABLE IF EXISTS `product_fields_x_card`;
CREATE TABLE IF NOT EXISTS `product_fields_x_card` (
  `counter` INTEGER UNSIGNED NOT NULL,
  `id_card` int(10) unsigned NOT NULL,
  `id_prod` int(10) unsigned NOT NULL,
  `id_field` int(10) unsigned NOT NULL,
  `qta_prod` int(10) unsigned NOT NULL,
  `value` varchar(250) NOT NULL,
  PRIMARY KEY  (`counter`,`id_card`,`id_prod`,`id_field`,`value`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `business_rules`
-- 
DROP TABLE IF EXISTS `business_rules`;
CREATE TABLE IF NOT EXISTS `business_rules` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `rule_type` int(10) unsigned NOT NULL,
  `label` varchar(100) NOT NULL,
  `description` text,
  `activate` smallint(1) unsigned NOT NULL default '0',
  `voucher_id` int(10) default NULL,
  PRIMARY KEY  (`id`),
  KEY `Index_label` (`label`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `business_rules_config`
-- 
DROP TABLE IF EXISTS `business_rules_config`;
CREATE TABLE IF NOT EXISTS `business_rules_config` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `id_rule` int(10) unsigned NOT NULL,
  `id_prod_orig` int(10) unsigned default NULL,
  `id_prod_ref` int(10) unsigned default NULL,
  `rate_from` decimal(10,2) NOT NULL,
  `rate_to` decimal(10,2) NOT NULL,
  `rate_from_ref` decimal(10,2) default NULL,
  `rate_to_ref` decimal(10,2) default NULL,
  `operation` smallint(1) unsigned NOT NULL default '0' COMMENT 'tipo di operatione da eseguire: 0 nulla, 1 somma, 2 sottrazione;',
  `applyto` smallint(1) unsigned NOT NULL default '0' COMMENT 'a chi deve essere applicata la regola: 0 nulla, 1 a prod_orig, 2 a prod_ref, 3 al meno caro, 4 al più caro, 5 a tutti e due;',
  `apply_4_qta` int(10) unsigned NOT NULL default '0' COMMENT 'per quale quantità  deve essere applicata la regola',
  `valore` decimal(10,2) NOT NULL,
  PRIMARY KEY  (`id`),
  UNIQUE KEY `Index_U` (`id_rule`,`id_prod_orig`,`rate_from`,`rate_to`,`id_prod_ref`),
  KEY `Index_From` (`rate_from`),
  KEY `Index_To` (`rate_to`),
  KEY `Index_Val` (`valore`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `business_rules_x_ordine`
-- 
DROP TABLE IF EXISTS `business_rules_x_ordine`;
CREATE TABLE IF NOT EXISTS `business_rules_x_ordine` (
  `id_rule` int(10) unsigned NOT NULL,
  `id_order` int(10) unsigned NOT NULL,
  `id_prod` int(10) unsigned NOT NULL default '0',
  `counter_prod` int(10) unsigned NOT NULL default '0',
  `label` VARCHAR(100) NOT NULL,
  `valore` DECIMAL(10,2) NOT NULL,   
  PRIMARY KEY (`id_rule`,`id_order`,`id_prod`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `voucher_campaign`
-- 
DROP TABLE IF EXISTS `voucher_campaign`;
CREATE TABLE IF NOT EXISTS `voucher_campaign` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `label` varchar(100) NOT NULL,
  `voucher_type` int(10) unsigned NOT NULL default '0' COMMENT 'tipo di voucher creato - 0=one shot, 1=per x volte, 2=one shot entro il periodo specificato, 3=per x volte entro il periodo specificato, 4=gift (come one shot ma con id_utente che ha fatto il regalo associato al voucher)',
  `description` text,
  `valore` decimal(10,2) NOT NULL,
  `operation` smallint(1) unsigned NOT NULL default '0' COMMENT 'tipo di calcolo applicato - 0=percentuale, 1=fisso',
  `activate` smallint(1) unsigned NOT NULL default '0',
  `max_generation` int(10) NOT NULL default '-1',
  `max_usage` int(10) NOT NULL default '-1',
  `enable_date` timestamp NULL default NULL,
  `expire_date` timestamp NULL default NULL,
  `exclude_prod_rule` smallint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (`id`),
  KEY `Index_label` (`label`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `voucher_code`
-- 
DROP TABLE IF EXISTS `voucher_code`;
CREATE TABLE IF NOT EXISTS `voucher_code` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `code` varchar(100) NOT NULL COMMENT 'il codice voucher verrà generato con un nuovo GUID ad hoc',
  `voucher_campaign` int(10) unsigned NOT NULL,
  `insert_date` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `usage_counter` int(10) unsigned NOT NULL default '0',
  `id_user_ref` int(10) unsigned default NULL,
  PRIMARY KEY  (`id`,`code`),
  KEY `Index_vc` (`voucher_campaign`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `voucher_x_ordine`
-- 
DROP TABLE IF EXISTS `voucher_x_ordine`;
CREATE TABLE IF NOT EXISTS `voucher_x_ordine` (
  `id_order` int(10) unsigned NOT NULL,
  `voucher_code` varchar(100) NOT NULL,
  `id_voucher` int(10) unsigned NOT NULL,
  `valore` decimal(10,2) NOT NULL,
  `insert_date` TIMESTAMP NOT NULL ,
  PRIMARY KEY  (`id_order`,`voucher_code`,`id_voucher`, `insert_date`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;