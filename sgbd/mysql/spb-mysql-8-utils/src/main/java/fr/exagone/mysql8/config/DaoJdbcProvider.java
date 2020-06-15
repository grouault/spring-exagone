package fr.exagone.mysql8.config;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.jdbc.core.JdbcTemplate;

import fr.exagone.mysql8.jdbc.repository.impl.ChatDAO;

@Configuration
public class DaoJdbcProvider {

	@Autowired
	JdbcTemplate jdbcTemplate;
	
	@Bean
	public ChatDAO getChatDao() {
		return new ChatDAO(jdbcTemplate);
	}
	
}
