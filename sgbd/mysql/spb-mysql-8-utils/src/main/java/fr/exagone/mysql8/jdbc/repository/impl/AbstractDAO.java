package fr.exagone.mysql8.jdbc.repository.impl;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.dao.EmptyResultDataAccessException;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.PreparedStatementCreator;
import org.springframework.jdbc.core.PreparedStatementSetter;
import org.springframework.jdbc.core.RowMapper;
import org.springframework.jdbc.support.GeneratedKeyHolder;
import org.springframework.jdbc.support.KeyHolder;

import fr.exagone.mysql8.jdbc.model.domain.IEntity;
import fr.exagone.mysql8.jdbc.repository.IDAO;
import fr.exagone.mysql8.jdbc.repository.ex.ExceptionDao;

public abstract class AbstractDAO<T extends IEntity> implements IDAO<T> {

	private static final Logger LOG = LogManager.getLogger();
	
	private JdbcTemplate jdbcTemplate;
	
	/**
	 * Constructeur de l'objet.
	 */
	protected AbstractDAO() {
		super();
	}
	
	protected AbstractDAO(JdbcTemplate jdbcTemplate) {
		this.jdbcTemplate = jdbcTemplate;
	}

	/**
	 * Retourne le nom de la table.
	 *
	 * @return le nom de la table.
	 */
	protected abstract String getTableName();

	/**
	 * Retourne la clef primaire de la table.
	 *
	 * @return la clef primaire de la table.
	 */
	protected String getPkName() {
		return "id";
	}

	/**
	 * Retourne la liste des colonnes de la table
	 *
	 * @return la liste des colonnes de la table
	 */
	protected abstract String getAllColumnNames();

	protected abstract RowMapper<T> getMapper();

	/**
	 * A la responsabilite de creer un statement qui servira pour l'insertion.
	 *
	 * @param pUneEntite
	 *            une entite a inserer
	 * @param connexion
	 *            une connexion
	 * @return un statement adapte a l'insertion
	 * @throws SQLException
	 *             si un probleme survient
	 */
	protected abstract PreparedStatement buildStatementForInsert(T pUneEntite, Connection connexion)
			throws SQLException;

	/**
	 * A la responsabilite de creer un statement qui servira pour la mise a
	 * jour.
	 *
	 * @param pUneEntite
	 *            une entite a mettre a jour
	 * @param connexion
	 *            une connexion
	 * @return un statement adapte a la mise a jour
	 * @throws SQLException
	 *             si un probleme survient
	 */
	protected abstract PreparedStatement buildStatementForUpdate(T pUneEntite, Connection connexion)
			throws SQLException;

	@Override
	public T insert(final T pUneEntite) throws ExceptionDao {
		if (pUneEntite == null) {
			return null;
		}
		AbstractDAO.LOG.debug("Insert {}", pUneEntite.getClass());
		try {
			
			// on passe par un prepare statement creator.
			PreparedStatementCreator psc = new PreparedStatementCreator() {
				@Override
				public PreparedStatement createPreparedStatement(Connection con) throws SQLException {
					return AbstractDAO.this.buildStatementForInsert(pUneEntite, con);
				}
			};
			
			// recuperation de l'identifiant.
			KeyHolder kh = new GeneratedKeyHolder();
			this.getJdbcTemplate().update(psc, kh);
			pUneEntite.setId(Integer.valueOf(kh.getKey().intValue()));
		
		} catch (Exception e) {
			throw new ExceptionDao(e);
		} 
		return pUneEntite;
	}

	@Override
	public T update(final T pUneEntite) throws ExceptionDao {
		
		if (pUneEntite == null) {
			return null;
		}
		
		AbstractDAO.LOG.debug("update {}", pUneEntite.getClass());
		if (pUneEntite.getId() == null) {
			throw new ExceptionDao("L'entite n'a pas d'ID");
		}

		try {

			// on passe par un prepare statement creator.
			PreparedStatementCreator psc = new PreparedStatementCreator() {
				@Override
				public PreparedStatement createPreparedStatement(Connection con) throws SQLException {
					return AbstractDAO.this.buildStatementForUpdate(pUneEntite, con);
				}
			};
			
			// execution de la mise à jour.
			if (this.getJdbcTemplate().update(psc) == 1) {
				return pUneEntite;
			}
			
		} catch (Exception e) {
			throw new ExceptionDao(e);
		} 
		
		return null;
		
	}

	@Override
	public boolean delete(T pUneEntite) throws ExceptionDao {
		if (pUneEntite == null) {
			return false;
		}
		AbstractDAO.LOG.debug("delete {}", pUneEntite.getClass());
		if (pUneEntite.getId() == null) {
			throw new ExceptionDao("L'entite n'a pas d'ID");
		}
		try {
			
			String toUpdate = "delete from " + this.getTableName() + " where " + this.getPkName() + "=?;";
			return 1 == this.getJdbcTemplate().update(toUpdate, pUneEntite.getId());
		} catch (Exception e) {
			throw new ExceptionDao(e);
		}
	}

	@Override
	public T select(int pUneClef) throws ExceptionDao {

		T result = null;
		AbstractDAO.LOG.debug("select sur {} avec id={}", this.getClass(), String.valueOf(pUneClef));
	
		try {
		
			// request
			String request = "select " + this.getAllColumnNames() + " from " + this.getTableName() + " where "
					+ this.getPkName() + "=?;";
			// result
			result = this.getJdbcTemplate().queryForObject(request, this.getMapper(), pUneClef);
		
		} catch (EmptyResultDataAccessException e) {
			return result;
		} catch (Exception e) {
			throw new ExceptionDao(e);
		} 

		return result;
	}

	@Override
	public List<T> selectAll(String pAWhere, String pAnOrderBy) throws ExceptionDao {
		
		List<T> result = new ArrayList<T>();
		AbstractDAO.LOG.debug("selectAll sur {} avec where={} prderBy={}", this.getClass(), pAWhere, pAnOrderBy);
		
		try {

			// request
			StringBuilder request = new StringBuilder();
			request.append("select ").append(this.getAllColumnNames()).append(" from ");
			request.append(this.getTableName());
			if (pAWhere != null) {
				request.append(" where ");
				request.append(pAWhere);
			}
			if (pAnOrderBy != null) {
				request.append(" order by ");
				request.append(pAnOrderBy);
			}
			request.append(';');
			AbstractDAO.LOG.debug("selectAll sur {} - requete={}", this.getClass(), request.toString());

			// execution de la query.
			result = this.jdbcTemplate.query(request.toString(), this.getMapper());

		} catch (Exception e) {
			throw new ExceptionDao(e);
		} 

		return result;
	}


	/**
	 * Place les éléments dans la requête.
	 * 
	 * @param gaps
	 * @return
	 */
	public PreparedStatementSetter setPssParameters(final List<?> gaps) {
		return new PreparedStatementSetter() {
			
			@Override
			public void setValues(PreparedStatement ps) throws SQLException {
				
				Iterator<?> iter = gaps.iterator();
				int id = 0;
				while (iter.hasNext()) {
					id++;
					Object lE = iter.next();
					if (lE == null) {
						continue;
					}
					if (lE instanceof String) {
						ps.setString(id, (String) lE);
					} else if (lE instanceof java.sql.Date) {
						ps.setDate(id, (java.sql.Date) lE);
					} else if (lE instanceof java.util.Date) {
						ps.setDate(id, new java.sql.Date(((java.util.Date) lE).getTime()));
					} else if (lE instanceof Timestamp) {
						ps.setTimestamp(id, (Timestamp) lE);
					} else if (lE instanceof Integer) {
						ps.setInt(id, ((Integer) lE).intValue());
					} else if (lE instanceof Double) {
						ps.setDouble(id, ((Double) lE).doubleValue());
					} else {
						throw new SQLException("Type invalid '" + lE.getClass().getSimpleName() + "'");
					}

				}
				
			}
			
		};		
	}
	
	
	public JdbcTemplate getJdbcTemplate() {
		return jdbcTemplate;
	}

	public void setJdbcTemplate(JdbcTemplate jdbcTemplate) {
		this.jdbcTemplate = jdbcTemplate;
	}
	
}
