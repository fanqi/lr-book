<%@ include file="/html/bacofis/init.jsp" %>

<% 
	ResultRow row = (ResultRow) request.getAttribute(WebKeys.SEARCH_CONTAINER_RESULT_ROW);
	Profile obj = (Profile) row.getObject();
%>

<%= obj.getProfileName() %>