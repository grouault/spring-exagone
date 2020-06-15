package fr.exagone.mysql8.jdbc.repository.impl;

import java.sql.Connection;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.time.LocalDate;

import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowMapper;

import fr.exagone.mysql8.jdbc.mappers.ChatMapper;
import fr.exagone.mysql8.jdbc.model.domain.impl.ChatEntity;

/**
 * Dao Jdbc pour l'entité Chat.
 * 
 * @author gildas
 *
 */
public class ChatDAO extends AbstractDAO<ChatEntity> {

	public ChatDAO(JdbcTemplate jdbcTemplate) {
		super(jdbcTemplate);
	}
	
	@Override
	protected String getTableName() {
		return "cat";
	}

	@Override
	protected String getAllColumnNames() {
		return "id, cat_name, birthday_date";
	}

	@Override
	protected RowMapper<ChatEntity> getMapper() {
		return new ChatMapper();
	}

	@Override
	protected PreparedStatement buildStatementForInsert(ChatEntity pUneEntite, Connection connexion)
			throws SQLException {
		
		String request = "insert into " + this.getTableName()
		+ " (cat_name, birthday_date) values (?,?);";
		PreparedStatement ps = connexion.prepareStatement(request, Statement.RETURN_GENERATED_KEYS);
		ps.setString(1, pUneEntite.getCatName());
		// cast sql date
		ps.setDate(2, Date.valueOf(pUneEntite.getBirthday()));
		return ps;
	
	}

	@Override
	protected PreparedStatement buildStatementForUpdate(ChatEntity pUneEntite, Connection connexion)
			throws SQLException {
		
		String request = "update " + this.getTableName()
		+ " set cat_name=?, birthday_date=? where " + this.getPkName() + "=?;";
		PreparedStatement ps = connexion.prepareStatement(request);
		ps.setString(1, pUneEntite.getCatName());
		// cast sql date
		ps.setDate(2, Date.valueOf(pUneEntite.getBirthday()));
		ps.setInt(3, pUneEntite.getId());
		return ps;
		
	}

}
