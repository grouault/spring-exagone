package fr.exagone.mysql8.jpa.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import fr.exagone.mysql8.jpa.domain.model.Chat;

@Repository
public interface ChatRepository  extends JpaRepository<Chat, Integer> {

}
