package fr.exagone.mysql8;

import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.time.LocalDate;
import java.util.Optional;

import org.junit.jupiter.api.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.orm.jpa.DataJpaTest;
import org.springframework.boot.test.autoconfigure.orm.jpa.TestEntityManager;
import org.springframework.test.context.junit4.SpringRunner;

import fr.exagone.mysql8.jpa.domain.model.Chat;
import fr.exagone.mysql8.jpa.repository.ChatRepository;

@RunWith(SpringRunner.class)
@DataJpaTest
public class ChatRepositoryJpaTest {
	
    @Autowired
    private TestEntityManager entityManager;
	
	@Autowired
	private ChatRepository chatRepository;
	
	@Test
	void whenFindById_thenTestNotNull() {
		
		Chat chat = new Chat();
		chat.setCatName("Max");
		chat.setBirthdayDate(LocalDate.of(2013, 06, 12));
		Chat persistEntity = entityManager.persist(chat);
		entityManager.flush();
		
		Optional<Chat> optChat = chatRepository.findById(persistEntity.getId());
		assertNotNull(optChat.get());
	}
	

}
