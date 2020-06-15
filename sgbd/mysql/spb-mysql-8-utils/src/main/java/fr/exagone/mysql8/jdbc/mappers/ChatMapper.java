package fr.exagone.mysql8.jdbc.mappers;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;

import org.springframework.jdbc.core.RowMapper;

import fr.exagone.mysql8.jdbc.model.domain.impl.ChatEntity;

public class ChatMapper implements RowMapper<ChatEntity>  {

	@Override
	public ChatEntity mapRow(ResultSet rs, int rowNum) throws SQLException {

		ChatEntity result = new ChatEntity();
		result.setId(Integer.valueOf(rs.getInt("id")));
		result.setCatName(rs.getString("cat_name"));
		// converted date
		Date dateToConvert = (Date)rs.getDate("birthday_date");
		// LocalDate convertedDate = ((Date)dateToConvert).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
		LocalDate convertedDate = new java.sql.Date(dateToConvert.getTime()).toLocalDate();
		result.setBirthday( convertedDate );

		return result;
	
	}

}
