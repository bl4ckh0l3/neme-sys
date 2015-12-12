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
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


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
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


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
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


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
  PRIMARY KEY  (`id_commento`),
  KEY `Index_2` (`id_element`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


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
  `tipo` char(1) NOT NULL,
  PRIMARY KEY  (`keyword`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `conferma_utente`
-- 
DROP TABLE IF EXISTS `conferma_utente`;
CREATE TABLE `conferma_utente` (
  `id_user` int(10) unsigned NOT NULL,
  `confirmation_code` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `file_x_news`
-- 
DROP TABLE IF EXISTS `file_x_news`;
CREATE TABLE IF NOT EXISTS `file_x_news` (
  `id_news` int(10) unsigned NOT NULL,
  `id_file` int(10) unsigned NOT NULL,
  KEY `Index_1` (`id_news`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `language`
-- 
DROP TABLE IF EXISTS `language`;
CREATE TABLE IF NOT EXISTS `language` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `descrizione` varchar(250) default NULL,
  `label` varchar(100) default NULL,
  `subdomain_active` int(1) unsigned NOT NULL default '0',
  `url_subdomain` varchar(250) default NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `language_disponibili`
-- 
DROP TABLE IF EXISTS `language_disponibili`;
CREATE TABLE `language_disponibili` (
  `keyword` VARCHAR(2) NOT NULL,
  `description` VARCHAR(45) NOT NULL,
  PRIMARY KEY (`keyword`)
) ENGINE = InnoDB  DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `multi_language`
-- 
DROP TABLE IF EXISTS `multi_language`;
CREATE TABLE IF NOT EXISTS `multi_language` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `keyword` varchar(100) NOT NULL,
  `IT` text,
  `GB` text,
  `FR` text,
  `DE` text,
  `SP` text,
  `PT` text,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


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
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


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
  PRIMARY KEY  (`id`),
  KEY `Index_2` (`titolo`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `newsletter`
-- 
DROP TABLE IF EXISTS `newsletter`;
CREATE TABLE `newsletter` (
  `id_newsletter` int(10) unsigned NOT NULL auto_increment,
  `descrizione` varchar(100) NOT NULL,
  `stato` int(10) unsigned NOT NULL default '0',
  `template` varchar(100) NOT NULL,
  PRIMARY KEY  (`id_newsletter`),
  KEY `Index_stato` (`stato`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


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
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `news_x_utente`
-- 
DROP TABLE IF EXISTS `news_x_utente`;
CREATE TABLE IF NOT EXISTS `news_x_utente` (
  `id_news` int(10) unsigned NOT NULL,
  `id_utente` int(10) unsigned NOT NULL,
  KEY `Index_1` (`id_news`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `ordini`
-- 
DROP TABLE IF EXISTS `ordini`;
CREATE TABLE IF NOT EXISTS `ordini` (
  `id_ordine` int(10) unsigned NOT NULL auto_increment,
  `id_utente` int(10) unsigned NOT NULL,
  `dta_inserimento` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `stato_ordine` varchar(100) NOT NULL,
  `totale_imponibile` DECIMAL(10,2) NOT NULL,
  `totale_tasse` DECIMAL(10,2) NOT NULL,
  `totale` decimal(10,2) NOT NULL,
  `tipo_pagam` varchar(100),
  `pagam_effettuato` int(10) unsigned NOT NULL,
  `order_guid` varchar(250) NOT NULL,
  `user_notified_x_download` INT(1) UNSIGNED NOT NULL DEFAULT '0',
  `notes` text,
  PRIMARY KEY  (`id_ordine`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


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
  `prezzo` double(10,2) NOT NULL,
  `qta_disp` varchar(100) NOT NULL,
  `attivo` int(10) unsigned NOT NULL,
  `sconto` varchar(2) NOT NULL,
  `codice_prod` varchar(100) NOT NULL,
  `id_tassa_applicata` int(10) unsigned default NULL,
  `downloadable` smallint(1) unsigned NOT NULL,
  `max_download` int(11) NOT NULL default '-1',
  `max_download_time` int(11) NOT NULL default '-1',
  PRIMARY KEY  (`id_prodotto`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;

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
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

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
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------
-- 
-- Struttura della tabella `prodotti_x_carrello`
-- 
DROP TABLE IF EXISTS `prodotti_x_carrello`;
CREATE TABLE IF NOT EXISTS `prodotti_x_carrello` (
  `id_carrello` int(10) unsigned NOT NULL,
  `id_prodotto` int(10) unsigned NOT NULL,
  `qta_prod` int(10) unsigned NOT NULL,
  PRIMARY KEY  (`id_carrello`,`id_prodotto`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `prodotti_x_ordine`
-- 
DROP TABLE IF EXISTS `prodotti_x_ordine`;
CREATE TABLE IF NOT EXISTS `prodotti_x_ordine` (
  `id_ordine` int(10) unsigned NOT NULL,
  `id_prodotto` int(10) unsigned NOT NULL,
  `nome_prodotto` varchar(100) NOT NULL,
  `qta` int(10) unsigned NOT NULL,
  `totale` double(10,2) NOT NULL,
  PRIMARY KEY  (`id_ordine`,`id_prodotto`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


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
  INDEX `Index_1`(`id_ordine`),
  INDEX `Index_2`(`id_spesa`)
)ENGINE = InnoDB  DEFAULT CHARSET=latin1;


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
  PRIMARY KEY  (`id`),
  KEY `descrizione` (`descrizione`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `target_type`
-- 
DROP TABLE IF EXISTS `target_type`;
CREATE TABLE IF NOT EXISTS `target_type` (
  `id` int(10) unsigned NOT NULL,
  `descrizione` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `target_x_categoria`
-- 
DROP TABLE IF EXISTS `target_x_categoria`;
CREATE TABLE IF NOT EXISTS `target_x_categoria` (
  `id_target` int(10) unsigned NOT NULL,
  `id_categoria` int(10) unsigned NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `target_x_news`
-- 
DROP TABLE IF EXISTS `target_x_news`;
CREATE TABLE IF NOT EXISTS `target_x_news` (
  `id_target` int(10) unsigned NOT NULL,
  `id_news` int(10) unsigned NOT NULL,
  KEY `Index_2` (`id_news`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `target_x_prodotto`
-- 
DROP TABLE IF EXISTS `target_x_prodotto`;
CREATE TABLE IF NOT EXISTS `target_x_prodotto` (
  `id_target` int(10) unsigned NOT NULL,
  `id_prodotto` int(10) unsigned NOT NULL,
  KEY `Index_1` (`id_prodotto`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `target_x_utente`
-- 
DROP TABLE IF EXISTS `target_x_utente`;
CREATE TABLE IF NOT EXISTS `target_x_utente` (
  `id_target` int(10) unsigned NOT NULL,
  `id_utente` int(10) unsigned NOT NULL,
  KEY `Index_1` USING BTREE (`id_utente`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `template_disponibili`
-- 
DROP TABLE IF EXISTS `template_disponibili`;
CREATE TABLE IF NOT EXISTS `template_disponibili` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `dir_template` varchar(50) NOT NULL,
  `template_css` varchar(100) default NULL,
  `descrizione` varchar(250) default NULL,
  `base_template` int(2) unsigned NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `page_x_template`
-- 
DROP TABLE IF EXISTS `page_x_template`;
CREATE TABLE `page_x_template` (
  `id` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
  `id_template` INTEGER UNSIGNED NOT NULL,
  `file_name` VARCHAR(50) NOT NULL,
  `page_num` INTEGER NOT NULL,
  PRIMARY KEY (`id`),
  INDEX `Index_2`(`id_template`),
  INDEX `Index_3`(`page_num`)
)ENGINE = InnoDB  DEFAULT CHARSET=latin1;


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
  `file_dida` varchar(250) default NULL,
  `file_label` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


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
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


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
  `file_dida` varchar(250) default NULL,
  `file_label` varchar(100) NOT NULL,
  `dta_ins` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `utenti`
-- 
DROP TABLE IF EXISTS `utenti`;
CREATE TABLE IF NOT EXISTS `utenti` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `username` varchar(100) NOT NULL,
  `pwd` varchar(50) NOT NULL,
  `nome` varchar(100) NOT NULL,
  `cognome` varchar(100) NOT NULL,
  `email` varchar(100) NOT NULL,
  `ruolo` int(10) unsigned NOT NULL,
  `privacy` char(1) NOT NULL,
  `newsletter` char(1) NOT NULL,
  `telephone` varchar(50) default NULL,
  `fax` varchar(50) default NULL,
  `companyName` varchar(100) default NULL,
  `address` varchar(250) default NULL,
  `city` varchar(100) default NULL,
  `zipCode` varchar(20) default NULL,
  `website` varchar(100) default NULL,
  `businessActivity` varchar(250) default NULL,
  `country` varchar(100) default NULL,
  `utenteAttivo` varchar(2) NOT NULL,
  `sconto` varchar(2) default NULL,
  `adminComments` text,
  `codFiscPiva` varchar(16) default NULL,
  `insertDate` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `modifyDate` timestamp NOT NULL default '0000-00-00 00:00:00',
  `birthday` varchar(100) default NULL,
  `sex` varchar(1) default NULL,
  `interests` varchar(250) default NULL,
  `list_others` varchar(250) default NULL,
  `public` INTEGER(1) UNSIGNED NOT NULL DEFAULT 0,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


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
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `ruoli_utente`
-- 
DROP TABLE IF EXISTS `ruoli_utente`;
CREATE TABLE IF NOT EXISTS `ruoli_utente` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `descrizione` varchar(100) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


-- --------------------------------------------------------
-- 
-- Struttura della tabella `spese_accessorie`
-- 
DROP TABLE IF EXISTS `spese_accessorie`;
CREATE TABLE `spese_accessorie` (
  `id` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
  `descrizione_spesa` VARCHAR(100) NOT NULL,
  `valore` INTEGER(11) UNSIGNED NOT NULL,
  `tipologia_valore` SMALLINT(1) UNSIGNED NOT NULL,
  `id_tassa_applicata` INTEGER(10) UNSIGNED default NULL,
  `applica_frontend`SMALLINT(1) UNSIGNED,
  `applica_backend` SMALLINT(1) UNSIGNED,
  PRIMARY KEY (`id`),
  INDEX `Index_2`(`valore`),
  INDEX `Index_3`(`id_tassa_applicata`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;


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
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;



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
  `url` smallint(1) unsigned NOT NULL default '0',
  `id_modulo` int(10) default NULL,
  `activate` smallint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

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
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

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
)ENGINE = InnoDB  DEFAULT CHARSET=latin1;

DROP TABLE IF EXISTS `payment_fixed_app_field`;
CREATE TABLE `payment_fixed_app_field` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `keyword` varchar(50) NOT NULL,
  `value` varchar(100) default NULL,
  `used` smallint(1) unsigned NOT NULL default '1',
  PRIMARY KEY  (`id`),
  UNIQUE KEY `Index_2` (`keyword`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

DROP TABLE IF EXISTS `paypal_field`;
CREATE TABLE `paypal_field` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `keyword` varchar(50) NOT NULL,
  `value` varchar(100) default NULL,
  `match_field` varchar(50) default NULL,
  PRIMARY KEY  (`id`),
  UNIQUE KEY `Index_2` (`keyword`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

DROP TABLE IF EXISTS `sella_field`;
CREATE TABLE `sella_field` (
  `id` int(10) unsigned NOT NULL auto_increment,
  `keyword` varchar(50) NOT NULL,
  `value` varchar(100) default NULL,
  `match_field` varchar(50) default NULL,
  PRIMARY KEY  (`id`),
  UNIQUE KEY `Index_2` (`keyword`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

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
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

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
) ENGINE=InnoDB  DEFAULT CHARSET=latin1;

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
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

DROP TABLE IF EXISTS `friend_x_user`;
CREATE TABLE `friend_x_user` (
  `id_friend` INTEGER UNSIGNED NOT NULL,
  `id_user` INTEGER UNSIGNED NOT NULL,
  PRIMARY KEY (`id_friend`, `id_user`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;