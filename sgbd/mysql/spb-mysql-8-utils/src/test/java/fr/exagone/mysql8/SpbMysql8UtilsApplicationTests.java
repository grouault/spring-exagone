package fr.exagone.mysql8;

import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.util.Optional;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import fr.exagone.mysql8.jpa.domain.model.Chat;
import fr.exagone.mysql8.jpa.repository.ChatRepository;

@SpringBootTest
class SpbMysql8UtilsApplicationTests {

	@Autowired
	ChatRepository chatRepository;
	
	@Test
	void contextLoads() {
		Optional<Chat> chat = chatRepository.findById(1);
		assertNotNull(chat.get());
	}
	

}
