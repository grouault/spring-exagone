package fr.exagone.mysql8.launcher.test;

import java.time.LocalDate;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import fr.exagone.mysql8.jdbc.model.domain.impl.ChatEntity;
import fr.exagone.mysql8.jdbc.repository.ex.ExceptionDao;
import fr.exagone.mysql8.jdbc.repository.impl.ChatDAO;

public class TestJdbc {

	private static Logger LOGGER = LogManager.getLogger();
	
	public static ChatEntity insertChat(ChatDAO daoChat) throws ExceptionDao {
		
		ChatEntity maxEntity = new ChatEntity();
		maxEntity.setBirthday(LocalDate.of(2013, 06, 12));
		maxEntity.setCatName("Max");
		
		LOGGER.info("Max Entity : {} ", maxEntity.toString());
		ChatEntity insertedChat = daoChat.insert(maxEntity);
		LOGGER.info("Insertion Entity : {}", maxEntity);
		
		return insertedChat;
		
	}
	
	
	public static ChatEntity updateChat (ChatDAO daoChat, Integer id) throws ExceptionDao {
		
		ChatEntity monChat = daoChat.select(id);
		LOGGER.info("[JDBC]-[Before-Update] - monChat = {}", monChat);
		monChat.setCatName("tutu");
		
		ChatEntity updatedChat = daoChat.update(monChat);
		LOGGER.info("[JDBC]-[After-Update] - updatedChat = {}", updatedChat);
	
		return updatedChat;
		
	}
	
}
