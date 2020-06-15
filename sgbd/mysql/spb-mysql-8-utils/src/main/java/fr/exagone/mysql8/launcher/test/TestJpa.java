package fr.exagone.mysql8.launcher.test;

import java.util.Optional;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import fr.exagone.mysql8.jpa.domain.model.Chat;
import fr.exagone.mysql8.jpa.repository.ChatRepository;

public class TestJpa {

	private static Logger LOGGER = LogManager.getLogger();
	
	public static Chat updateChat(ChatRepository chatRepo) {
		
		Optional<Chat> chat = chatRepo.findById(1);
		Chat leChat = chat.get();
		leChat.setCatName("totor");
		Chat toReturn = chatRepo.save(leChat);
		LOGGER.info("[JPA]-[After-Update] - updatedChat = {}", toReturn);
		
		return toReturn;
		
	}
	
}
