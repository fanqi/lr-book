/**
 * Copyright (c) 2000-2013 Liferay, Inc. All rights reserved.
 *
 * This library is free software; you can redistribute it and/or modify it under
 * the terms of the GNU Lesser General Public License as published by the Free
 * Software Foundation; either version 2.1 of the License, or (at your option)
 * any later version.
 *
 * This library is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU Lesser General Public License for more
 * details.
 */

package com.inikah.slayer.service.impl;

import java.util.List;

import com.inikah.slayer.model.MatchCriteria;
import com.inikah.slayer.model.Profile;
import com.inikah.slayer.service.base.MatchCriteriaLocalServiceBaseImpl;
import com.inikah.util.AgeUtil;
import com.liferay.portal.kernel.dao.orm.DynamicQuery;
import com.liferay.portal.kernel.dao.orm.DynamicQueryFactoryUtil;
import com.liferay.portal.kernel.dao.orm.RestrictionsFactoryUtil;
import com.liferay.portal.kernel.exception.SystemException;
import com.liferay.portal.kernel.portlet.PortletClassLoaderUtil;
import com.liferay.portal.kernel.util.Validator;

/**
 * The implementation of the match criteria local service.
 *
 * <p>
 * All custom service methods should be put in this class. Whenever methods are added, rerun ServiceBuilder to copy their definitions into the {@link com.inikah.slayer.service.MatchCriteriaLocalService} interface.
 *
 * <p>
 * This is a local service. Methods of this service will not have security checks based on the propagated JAAS credentials because this service can only be accessed from within the same VM.
 * </p>
 *
 * @author Ahmed Hasan
 * @see com.inikah.slayer.service.base.MatchCriteriaLocalServiceBaseImpl
 * @see com.inikah.slayer.service.MatchCriteriaLocalServiceUtil
 */
public class MatchCriteriaLocalServiceImpl
	extends MatchCriteriaLocalServiceBaseImpl {
	/*
	 * NOTE FOR DEVELOPERS:
	 *
	 * Never reference this interface directly. Always use {@link com.inikah.slayer.service.MatchCriteriaLocalServiceUtil} to access the match criteria local service.
	 */
	
	public void init(Profile profile) {
		
		MatchCriteria matchCriteria = createMatchCriteria(profile.getProfileId());
		
		double age = AgeUtil.getComputeAge(profile.getBornOn());
		
		int minAge = 0;
		int maxAge = 0;
		
		matchCriteria.setGroom(profile.getBride());
		
		if (profile.getBride()) {
			minAge = (int) age - 2;
			maxAge = minAge + 12;
		} else {
			minAge = (int) age - 10;
			maxAge = minAge + 12;
		}
		
		// set Location
		
		matchCriteria.setMinAge(minAge);
		matchCriteria.setMaxAge(maxAge);
		
		try {
			addMatchCriteria(matchCriteria);
		} catch (SystemException e) {
 			e.printStackTrace();
		}
	}
	
	public List<Profile> getMatches(long profileId) {
		
		List<Profile> matches = null;
		
		MatchCriteria matchCriteria = null;
		try {
			matchCriteria = fetchMatchCriteria(profileId);
		} catch (SystemException e) {
			e.printStackTrace();
		}
		
		if (Validator.isNull(matchCriteria)) {
			return null;
		}

		DynamicQuery dynamicQuery = DynamicQueryFactoryUtil.forClass(Profile.class, PortletClassLoaderUtil.getClassLoader());
		
		// bride or grooms
		dynamicQuery.add(RestrictionsFactoryUtil.eq("groom", matchCriteria.getGroom()));
		
		// minAge and maxAge
		int minValue = AgeUtil.getBornOnFigure(matchCriteria.getMinAge(), AgeUtil.MIN);
		int maxValue = AgeUtil.getBornOnFigure(matchCriteria.getMaxAge(), AgeUtil.MAX);
		dynamicQuery.add(RestrictionsFactoryUtil.between("bornOn", minValue, maxValue));
		
		// exclude the profiles of the same user
		dynamicQuery.add(RestrictionsFactoryUtil.ne("userId", matchCriteria.getUserId()));
		
		return matches;
		
	}
}