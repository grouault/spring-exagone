package fr.exagone.mysql8.jdbc.repository;

import java.util.List;

import fr.exagone.mysql8.jdbc.model.domain.IEntity;
import fr.exagone.mysql8.jdbc.repository.ex.ExceptionDao;

/**
 * Interface mere de tous les DAO. <br/>
 * Ne pas oublier de changer le mot de passe si necessaire.
 *
 * @param <T>
 *            Un type d'entite.
 */
public interface IDAO<T extends IEntity> {

	/**
	 * Ajoute un element dans la base.
	 *
	 * @param pUneEntite
	 *            un element
	 * @param connexion
	 *            une connection
	 * @return l'objet ajoute
	 * @throws ExceptionDao
	 *             en cas de probleme
	 */
	public abstract T insert(T pUneEntite) throws ExceptionDao;

	/**
	 * Met un element a jour dans la base.
	 *
	 * @param pUneEntite
	 *            un element
	 * @param connexion
	 *            une connection
	 * @return l'objet mis a jour
	 * @throws ExceptionDao
	 *             en cas de probleme
	 */
	public abstract T update(T pUneEntite) throws ExceptionDao;

	/**
	 * Suprime un compte de la base.
	 *
	 * @param pUneEntite
	 *            un objet compte
	 * @return true si il a ete supprime, false sinon
	 * @throws ExceptionDao
	 *             en cas de probleme
	 */
	public abstract boolean delete(T pUneEntite) throws ExceptionDao;

	/**
	 * Selectionne un compte dans la base en fonction de sa clef.
	 *
	 * @param pUneClef
	 *            une clef
	 * @param connexion
	 *            une connection
	 * @return l'objet compte trouve, null sinon
	 * @throws ExceptionDao
	 *             en cas de probleme
	 */
	public abstract T select(int pUneClef) throws ExceptionDao;

	/**
	 * Selectionne tous les comptes dans la base.
	 *
	 * @param pAWhere
	 *            une condition where (sans le mot clef where)
	 * @param pAnOrderBy
	 *            une condition d'order by (sans le mot clef order by)
	 * @param connexion
	 *            une connection
	 * @return la liste des comptes trouves, une liste vide sinon
	 * @throws ExceptionDao
	 *             en cas de probleme
	 */
	public abstract List<T> selectAll(String pAWhere, String pAnOrderBy) throws ExceptionDao;


}