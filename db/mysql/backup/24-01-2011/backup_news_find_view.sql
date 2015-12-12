--
-- Definition of view `news_find`
--

DROP TABLE IF EXISTS `news_find`;
DROP VIEW IF EXISTS `news_find`;
##CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `news_find` AS select `news`.`id` AS `id`,`news`.`titolo` AS `titolo`,`news`.`abstract` AS `abstract`,`news`.`abstract_2` AS `abstract_2`,`news`.`abstract_3` AS `abstract_3`,`news`.`testo` AS `testo`,`news`.`keyword` AS `keyword`,`news`.`data_inserimento` AS `data_inserimento`,`news`.`data_pubblicazione` AS `data_pubblicazione`,`news`.`data_cancellazione` AS `data_cancellazione`,`news`.`stato_news` AS `stato_news`,`news_x_utente`.`id_utente` AS `id_utente` from (`news_x_utente` join `news` on((`news_x_utente`.`id_news` = `news`.`id`)));
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `news_find` AS select `news`.`id` AS `id`,`news`.`titolo` AS `titolo`,`news`.`abstract` AS `abstract`,`news`.`abstract_2` AS `abstract_2`,`news`.`abstract_3` AS `abstract_3`,`news`.`testo` AS `testo`,`news`.`keyword` AS `keyword`,`news`.`data_inserimento` AS `data_inserimento`,`news`.`data_pubblicazione` AS `data_pubblicazione`,`news`.`data_cancellazione` AS `data_cancellazione`,`news`.`stato_news` AS `stato_news`,`news_x_utente`.`id_utente` AS `id_utente` from (`news_x_utente` join `news` on((`news_x_utente`.`id_news` = `news`.`id`)));