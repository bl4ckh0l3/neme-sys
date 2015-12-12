--
-- Definition of constraints
--

ALTER TABLE `attach_x_prodotti` ADD FOREIGN KEY (`id_prodotto`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
ALTER TABLE `commenti_prodotto` ADD FOREIGN KEY (`id_prodotto`) REFERENCES `prodotti` (`id_prodotto`) ON DELETE CASCADE;
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