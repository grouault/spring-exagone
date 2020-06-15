package fr.exagone.mysql8;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;

import fr.exagone.mysql8.jdbc.model.domain.impl.ChatEntity;
import fr.exagone.mysql8.jdbc.repository.ex.ExceptionDao;
import fr.exagone.mysql8.jdbc.repository.impl.ChatDAO;
import fr.exagone.mysql8.jpa.repository.ChatRepository;
import fr.exagone.mysql8.launcher.test.TestJdbc;
import fr.exagone.mysql8.launcher.test.TestJpa;
import fr.exagone.mysql8.launcher.test.TestTimeZone;

@SpringBootApplication
public class SpbMysql8UtilsApplication {

	public static void main(String[] args) {
		
		Logger LOGGER = LogManager.getLogger();
		
		ApplicationContext spring = SpringApplication.run(SpbMysql8UtilsApplication.class, args);
		
		// JDBC
		ChatDAO daoChat = spring.getBean(ChatDAO.class);

		///////////////////
		// test time-zone
		///////////////////
		TestTimeZone.test1();

		//////////////
		// TEST JDBC
		//////////////
		try {
			// ChatEntity chat = TestJdbc.insertChat(daoChat);
			ChatEntity chat = daoChat.select(1);
			TestJdbc.updateChat(daoChat, chat.getId());
		} catch (ExceptionDao e) {
			LOGGER.error("erreur : {0}", e.getMessage());
		}
		
		//////////////
		// TEST JPA //
		//////////////
		ChatRepository chatRepo = spring.getBean(ChatRepository.class);
		TestJpa.updateChat(chatRepo);
		
	}

}
