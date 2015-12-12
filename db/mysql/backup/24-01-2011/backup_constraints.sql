--
-- Definition of constraints
--

ALTER TABLE `attach_x_prodotti` ADD FOREIGN KEY (`id_prodotto`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `file_x_news` ADD FOREIGN KEY (`id_news`) REFERENCES `news` (`id`) ON DELETE CASCADE;
ALTER TABLE `news_x_utente` ADD FOREIGN KEY (`id_news`) REFERENCES `news` (`id`) ON DELETE CASCADE;
ALTER TABLE `prodotti_x_carrello` ADD FOREIGN KEY (`id_carrello`) REFERENCES `carrello` (`id_carrello`) ON DELETE CASCADE;
ALTER TABLE `prodotti_x_ordine` ADD FOREIGN KEY (`id_ordine`) REFERENCES `ordini` (`id_ordine`) ON DELETE CASCADE;
ALTER TABLE `target_x_news` ADD FOREIGN KEY (`id_news`) REFERENCES `news` (`id`) ON DELETE CASCADE;
ALTER TABLE `target_x_prodotto` ADD FOREIGN KEY (`id_prodotto`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `target_x_utente` ADD FOREIGN KEY (`id_utente`) REFERENCES `utenti` (`id`) ON DELETE CASCADE;
ALTER TABLE `spese_x_ordine` ADD FOREIGN KEY (`id_ordine`) REFERENCES `ordini` (`id_ordine`) ON DELETE CASCADE;
ALTER TABLE `target_x_categoria` ADD FOREIGN KEY (`id_categoria`) REFERENCES `categorie` (`id`) ON DELETE CASCADE;
ALTER TABLE `attach_x_prodotti` ADD FOREIGN KEY (`id_prodotto`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `utenti_images` ADD FOREIGN KEY (`id_utente`) REFERENCES `utenti` (`id`) ON DELETE CASCADE;
ALTER TABLE `user_preference` ADD FOREIGN KEY (`id_user`) REFERENCES `utenti` (`id`) ON DELETE CASCADE;
ALTER TABLE `user_preference` ADD FOREIGN KEY (`id_friend`) REFERENCES `utenti` (`id`);
ALTER TABLE `friend_x_user` ADD FOREIGN KEY (`id_friend`) REFERENCES `utenti` (`id`) ON DELETE CASCADE;
ALTER TABLE `friend_x_user` ADD FOREIGN KEY (`id_user`) REFERENCES `utenti` (`id`) ON DELETE CASCADE;
ALTER TABLE `shipping_address` ADD FOREIGN KEY (`id_user`) REFERENCES `utenti` (`id`) ON DELETE CASCADE;
ALTER TABLE `order_shipping_address` ADD FOREIGN KEY (`id_order`) REFERENCES `ordini` (`id_ordine`) ON DELETE CASCADE;
ALTER TABLE `order_shipping_address` ADD FOREIGN KEY (`id_shipping`) REFERENCES `shipping_address` (`id`) ON DELETE CASCADE;
ALTER TABLE `downloadable_products` ADD FOREIGN KEY (`id_product`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `usr_group_x_margin_disc` ADD FOREIGN KEY (`id_marg_disc`) REFERENCES `margin_discount` (`id`) ON DELETE CASCADE;
ALTER TABLE `usr_group_x_margin_disc` ADD FOREIGN KEY (`id_user_group`) REFERENCES `user_group` (`id`) ON DELETE CASCADE;
ALTER TABLE `user_fields_match` ADD FOREIGN KEY (`id_user`) REFERENCES `utenti` (`id`) ON DELETE CASCADE;
ALTER TABLE `user_fields_match` ADD FOREIGN KEY (`id_field`) REFERENCES `user_fields` (`id`) ON DELETE CASCADE;
ALTER TABLE `user_fields` ADD FOREIGN KEY (`id_group`) REFERENCES `user_fields_group` (`id`);
ALTER TABLE `user_fields` ADD FOREIGN KEY (`type`) REFERENCES `user_fields_type` (`id`);
ALTER TABLE `product_fields_values` ADD FOREIGN KEY (`id_field`) REFERENCES `product_fields` (`id`) ON DELETE CASCADE;
ALTER TABLE `product_fields_match` ADD FOREIGN KEY (`id_prod`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `product_fields_match` ADD FOREIGN KEY (`id_field`) REFERENCES `product_fields` (`id`) ON DELETE CASCADE;
ALTER TABLE `product_fields_value_match` ADD FOREIGN KEY (`id_prod`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `product_fields_value_match` ADD FOREIGN KEY (`id_field`) REFERENCES `product_fields` (`id`) ON DELETE CASCADE;
ALTER TABLE `product_fields_x_order` ADD FOREIGN KEY (`id_order`) REFERENCES `ordini` (`id_ordine`) ON DELETE CASCADE;
ALTER TABLE `product_fields` ADD FOREIGN KEY (`id_group`) REFERENCES `product_fields_group` (`id`);
ALTER TABLE `product_fields` ADD FOREIGN KEY (`type`) REFERENCES `product_fields_type` (`id`);
ALTER TABLE `product_fields_x_card` ADD FOREIGN KEY (`id_card`) REFERENCES `carrello` (`id_carrello`) ON DELETE CASCADE;
ALTER TABLE `relation_x_prodotto` ADD FOREIGN KEY (`id_prod`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `relation_x_prodotto` ADD FOREIGN KEY (`id_prod_ref`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;