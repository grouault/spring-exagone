package fr.exagone.mysql8.jdbc.model.domain.impl;

import java.time.LocalDate;

public class ChatEntity extends AbstractEntity {

	private static final long serialVersionUID = 1L;

	private String catName;
	private LocalDate birthday;
	
	public ChatEntity() {
		// TODO Auto-generated constructor stub
	}
	
	public ChatEntity(Integer id, String catName, LocalDate birthday) {
		super(id);
		this.catName = catName;
		this.birthday = birthday;
	}
	
	public String getCatName() {
		return catName;
	}

	public void setCatName(String catName) {
		this.catName = catName;
	}

	public LocalDate getBirthday() {
		return birthday;
	}

	public void setBirthday(LocalDate birthday) {
		this.birthday = birthday;
	}

	@Override
	protected String asString() {
		StringBuilder sb = new StringBuilder();
		sb.append("catName=");
		sb.append(this.getCatName());
		sb.append(",birthday=");
		sb.append(this.getBirthday());
		return sb.toString();
	}

}
