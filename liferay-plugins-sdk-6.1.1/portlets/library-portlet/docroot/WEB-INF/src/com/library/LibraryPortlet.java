package com.library;

import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Date;
import java.util.List;

import javax.activation.MimetypesFileTypeMap;
import javax.portlet.ActionRequest;
import javax.portlet.ActionResponse;
import javax.portlet.Event;
import javax.portlet.EventRequest;
import javax.portlet.EventResponse;
import javax.portlet.PortletConfig;
import javax.portlet.PortletException;
import javax.portlet.PortletMode;
import javax.portlet.PortletPreferences;
import javax.portlet.PortletRequest;
import javax.portlet.PortletURL;
import javax.portlet.ProcessEvent;
import javax.portlet.RenderRequest;
import javax.portlet.RenderResponse;
import javax.portlet.ResourceRequest;
import javax.portlet.ResourceResponse;
import javax.servlet.http.HttpServletResponse;

import com.liferay.counter.service.CounterLocalServiceUtil;
import com.liferay.portal.kernel.exception.PortalException;
import com.liferay.portal.kernel.exception.SystemException;
import com.liferay.portal.kernel.json.JSONArray;
import com.liferay.portal.kernel.json.JSONException;
import com.liferay.portal.kernel.json.JSONFactoryUtil;
import com.liferay.portal.kernel.json.JSONObject;
import com.liferay.portal.kernel.servlet.ServletResponseUtil;
import com.liferay.portal.kernel.servlet.SessionErrors;
import com.liferay.portal.kernel.upload.UploadPortletRequest;
import com.liferay.portal.kernel.util.Constants;
import com.liferay.portal.kernel.util.FileUtil;
import com.liferay.portal.kernel.util.GetterUtil;
import com.liferay.portal.kernel.util.ParamUtil;
import com.liferay.portal.kernel.util.PortalClassInvoker;
import com.liferay.portal.kernel.util.StringPool;
import com.liferay.portal.kernel.util.Validator;
import com.liferay.portal.kernel.util.WebKeys;
import com.liferay.portal.model.Address;
import com.liferay.portal.model.CompanyConstants;
import com.liferay.portal.model.Contact;
import com.liferay.portal.model.Image;
import com.liferay.portal.model.ListType;
import com.liferay.portal.model.ResourceConstants;
import com.liferay.portal.security.permission.ActionKeys;
import com.liferay.portal.service.AddressLocalServiceUtil;
import com.liferay.portal.service.ImageLocalServiceUtil;
import com.liferay.portal.service.ListTypeServiceUtil;
import com.liferay.portal.service.ResourceLocalServiceUtil;
import com.liferay.portal.service.ServiceContext;
import com.liferay.portal.service.ServiceContextFactory;
import com.liferay.portal.service.permission.PortletPermissionUtil;
import com.liferay.portal.theme.ThemeDisplay;
import com.liferay.portal.util.PortalUtil;
import com.liferay.portlet.PortletURLFactoryUtil;
import com.liferay.portlet.asset.service.AssetEntryLocalServiceUtil;
import com.liferay.portlet.documentlibrary.store.DLStoreUtil;
import com.liferay.util.bridges.mvc.MVCPortlet;
import com.liferay.util.portlet.PortletProps;
import com.slayer.model.LMSBook;
import com.slayer.service.LMSBookLocalServiceUtil;
import com.util.LMSUtil;

/**
 * Portlet implementation class LibraryPortlet
 */
public class LibraryPortlet extends MVCPortlet {
	
	public void updateBook(ActionRequest actionRequest,
			ActionResponse actionResponse) throws IOException, PortletException {

		String bookTitle = ParamUtil.getString(actionRequest, "bookTitle");
		String author = ParamUtil.getString(actionRequest, "author");

		long bookId = ParamUtil.getLong(actionRequest, "bookId");

		ServiceContext serviceContext = null;
		try {
			serviceContext = ServiceContextFactory.getInstance(
					LMSBook.class.getName(), actionRequest);
		} catch (PortalException e) {
			e.printStackTrace();
		} catch (SystemException e) {
			e.printStackTrace();
		}

		LMSBook lmsBook = (bookId > 0l) ? LMSBookLocalServiceUtil.modifyBook(
				bookId, bookTitle, author) : LMSBookLocalServiceUtil
				.insertBook(bookTitle, author, serviceContext);
		saveAddress(actionRequest, lmsBook);

		// redirect after insert
		ThemeDisplay themeDisplay = (ThemeDisplay) actionRequest
				.getAttribute(WebKeys.THEME_DISPLAY);

		PortletConfig portletConfig = (PortletConfig) actionRequest
				.getAttribute("javax.portlet.config");

		String portletName = portletConfig.getPortletName();

		PortletURL successPageURL = PortletURLFactoryUtil.create(actionRequest,
				portletName + "_WAR_" + portletName + "portlet",
				themeDisplay.getPlid(), PortletRequest.RENDER_PHASE);

		successPageURL.setParameter("jspPage",
				(bookId > 0l) ? LibraryConstants.PAGE_LIST
						: LibraryConstants.PAGE_SUCCESS);
		actionResponse.sendRedirect(successPageURL.toString());
	}
	
	private void saveAddress(PortletRequest request, LMSBook lmsBook) {
		boolean saveAddress = ParamUtil.getBoolean(request, "saveAddress",
				false);

		if (!saveAddress)
			return;

		long addressId = ParamUtil.getLong(request, "addressId", 0l);
		long countryId = ParamUtil.getLong(request, "countryId");
		String city = ParamUtil.getString(request, "city");
		String street1 = ParamUtil.getString(request, "street1");
		String street2 = ParamUtil.getString(request, "street2");
		String zip = ParamUtil.getString(request, "zip");

		long userId = PortalUtil.getUserId(request);
		long companyId = PortalUtil.getCompanyId(request);

		saveAddress(addressId, countryId, city, street1, street2, zip, userId,
				companyId, lmsBook.getBookId(), LMSBook.class.getName());
	}

	private void saveAddress(long addressId, long countryId, String city,
			String street1, String street2, String zip, long userId,
			long companyId, long parentId, String parentClassName) {

		// 1. creating the address object (fresh or old)
		Address address = getAddress(addressId);

		// 2. setting the UI fields
		address.setCountryId(countryId);
		address.setCity(city);
		address.setStreet1(street1);
		address.setStreet2(street2);
		address.setZip(zip);

		// 3. set audit fields
		address.setUserId(userId);
		address.setCompanyId(companyId);

		// 4. Set address type and if primary.
		address.setTypeId(getTypeId("address", "personal"));
		address.setPrimary(true);

		// 5. most importantly set the parent details.
		address.setClassPK(parentId);
		address.setClassName(parentClassName);

		// 6. finally update the object (save address)
		try {
			AddressLocalServiceUtil.updateAddress(address);
		} catch (SystemException e) {
			e.printStackTrace();
		}
	}

	private int getTypeId(String entity, String type) {

		int typeId = 0;
		List<ListType> listTypes = null;

		try {
			listTypes = ListTypeServiceUtil.getListTypes(Contact.class
					.getName() + "." + entity);
		} catch (SystemException e) {
			e.printStackTrace();
		}

		for (ListType listType : listTypes) {
			if (listType.getName().equalsIgnoreCase(type)) {
				typeId = listType.getListTypeId();
				break;
			}
		}

		return typeId;
	}

	private Address getAddress(long addressId) {
		Address address = null;
		if (addressId > 0l) {
			try {
				address = AddressLocalServiceUtil.fetchAddress(addressId);
				address.setModifiedDate(new Date());
			} catch (SystemException e) {
				e.printStackTrace();
			}
		} else {
			try {
				addressId = CounterLocalServiceUtil.increment();
			} catch (SystemException e) {
				e.printStackTrace();
			}
			address = AddressLocalServiceUtil.createAddress(addressId);
			address.setCreateDate(new Date());
		}
		return address;
	}
	
	public void deleteBook(ActionRequest actionRequest,
			ActionResponse actionResponse) throws IOException, PortletException {

		long bookId = ParamUtil.getLong(actionRequest, "bookId");

		if (bookId > 0l) { // valid bookId
			try {
				LMSBookLocalServiceUtil.deleteLMSBook(bookId);
			} catch (PortalException e) {
				e.printStackTrace();
			} catch (SystemException e) {
				e.printStackTrace();
			}

			// delete the resource
			LMSBook lmsBook = null;
			try {
				lmsBook = LMSBookLocalServiceUtil.fetchLMSBook(bookId);
			} catch (SystemException e) {
				e.printStackTrace();
			}
			try {
				ResourceLocalServiceUtil.deleteResource(lmsBook,
						ResourceConstants.SCOPE_INDIVIDUAL);
			} catch (PortalException e) {
				e.printStackTrace();
			} catch (SystemException e) {
				e.printStackTrace();
			}

			// delete the asset
			try {
				AssetEntryLocalServiceUtil.deleteEntry(LMSBook.class.getName(),
						lmsBook.getPrimaryKey());
			} catch (PortalException e) {
				e.printStackTrace();
			} catch (SystemException e) {
				e.printStackTrace();
			}
		}

		// gracefully redirecting to the default list view
		String redirectURL = ParamUtil.getString(actionRequest, "redirectURL");
		actionResponse.sendRedirect(redirectURL);
	}

	public void render(RenderRequest request, RenderResponse response)
			throws PortletException, IOException {

		setSortParams(request);
		checkBeforeServe(request);
		super.render(request, response);
	}

	private void checkBeforeServe(RenderRequest request)
			throws PortletException {

		String jspPage = ParamUtil.getString(request, "jspPage");

		if (jspPage.equalsIgnoreCase(LibraryConstants.PAGE_UPDATE)) {
			ThemeDisplay themeDisplay = (ThemeDisplay) request
					.getAttribute(WebKeys.THEME_DISPLAY);

			StringBuilder portletName = new StringBuilder()
					.append(getPortletName()).append("_WAR_")
					.append(getPortletName()).append("portlet");

			try {
				PortletPermissionUtil.check(
						themeDisplay.getPermissionChecker(),
						portletName.toString(), ActionKeys.ADD_ENTRY);
			} catch (PortalException e) {
				e.printStackTrace();
			} catch (SystemException e) {
				e.printStackTrace();
			}
		}
	}

	private void setSortParams(RenderRequest request) {
		String jspPage = ParamUtil.getString(request, "jspPage");

		if (jspPage.equalsIgnoreCase(LibraryConstants.PAGE_LIST)) {
			String orderByCol = ParamUtil.getString(request, "orderByCol",
					"bookTitle");
			request.setAttribute("orderByCol", orderByCol);

			String orderByType = ParamUtil.getString(request, "orderByType",
					"asc");
			request.setAttribute("orderByType", orderByType);
		}
	}

	public void searchBooks(ActionRequest actionRequest,
			ActionResponse actionResponse) throws IOException, PortletException {

		String searchTerm = ParamUtil.getString(actionRequest, "searchTerm");
		
		if (Validator.isNull(searchTerm)) return;

		ThemeDisplay themeDisplay = (ThemeDisplay) 
			actionRequest.getAttribute(WebKeys.THEME_DISPLAY);
		
		try {
			List<LMSBook> lmsBooks = LMSBookLocalServiceUtil
				.searchIndex(searchTerm, 
					themeDisplay.getCompanyId(), themeDisplay.getScopeGroupId());

			actionRequest.setAttribute("SEARCH_RESULT", lmsBooks);
			actionResponse.setRenderParameter("jspPage",
					LibraryConstants.PAGE_LIST);

		} catch (SystemException e) {
			e.printStackTrace();
		}
	}

	public void deleteBooks(ActionRequest actionRequest,
			ActionResponse actionResponse) throws IOException, PortletException {

		String bookIdsForDelete = ParamUtil.getString(actionRequest,
				"bookIdsForDelete");

		// convert this into JSON format.
		bookIdsForDelete = "[" + bookIdsForDelete + "]";

		// The presence of ":" in the string
		// creates problem while parsing.
		// replace all occurance of ":" with
		// some other unique string, eg. "-"
		bookIdsForDelete = bookIdsForDelete.replaceAll(":", "-");

		// parse and get a JSON array of objects
		JSONArray jsonArray = null;
		try {
			jsonArray = JSONFactoryUtil.createJSONArray(bookIdsForDelete);
		} catch (JSONException e) {
			e.printStackTrace();
		}

		// process the jsonArray
		if (Validator.isNotNull(jsonArray)) {
			for (int i = 0; i < jsonArray.length(); i++) {
				JSONObject jsonObject = jsonArray.getJSONObject(i);

				long bookId = jsonObject.getLong("bookId");
				try {
					LMSBookLocalServiceUtil.deleteLMSBook(bookId);
				} catch (PortalException e) {
					e.printStackTrace();
				} catch (SystemException e) {
					e.printStackTrace();
				}
			}
		}

		// redirect to the list page again.
		actionResponse
				.setRenderParameter("jspPage", LibraryConstants.PAGE_LIST);
	}

	public void setPreferences(ActionRequest actionRequest,
			ActionResponse actionResponse) throws IOException, PortletException {

		String maxBooksLimit = ParamUtil.getString(actionRequest,
				"maxBooksLimit");

		PortletPreferences preferences = actionRequest.getPreferences();
		preferences.setValue("maxBooksLimit", maxBooksLimit);
		preferences.store();

		actionResponse.setPortletMode(PortletMode.VIEW);
	}


	@ProcessEvent(qname = "{http://liferay.com}lmsBook")
	public void quickAdd(EventRequest request, EventResponse response)
			throws PortletException, IOException {

		Event event = request.getEvent();

		LMSBook lmsBook = (LMSBook) event.getValue();

		ServiceContext serviceContext = null;
		try {
			serviceContext = ServiceContextFactory.getInstance(request);
		} catch (PortalException e) {
			e.printStackTrace();
		} catch (SystemException e) {
			e.printStackTrace();
		}

		LMSBookLocalServiceUtil.insertBook(lmsBook.getBookTitle(),
				lmsBook.getAuthor(), serviceContext);

		response.setRenderParameter("jspPage", LibraryConstants.PAGE_LIST);
	}

	public void UploadFiles(ActionRequest actionRequest,
			ActionResponse actionResponse) throws IOException, PortletException {

		UploadPortletRequest uploadRequest = PortalUtil
				.getUploadPortletRequest(actionRequest);

		ServiceContext serviceContext = null;
		try {
			serviceContext = ServiceContextFactory.getInstance(actionRequest);
		} catch (PortalException e) {
			e.printStackTrace();
		} catch (SystemException e) {
			e.printStackTrace();
		}

		File coverImage = uploadRequest.getFile("coverImage");

		long coverImageSize = coverImage.getTotalSpace();

		if (coverImageSize > 0) {
			long fileSize = FileUtil.getBytes(coverImage).length;

			long sizeLimit = GetterUtil.getLong(PortletProps
					.get(LibraryConstants.PROP_BOOK_COVER_IMAGE_MAX_SIZE));

			if (fileSize <= sizeLimit) {
				serviceContext.setAttribute("COVER_IMAGE", coverImage);
			} else {
				SessionErrors.add(actionRequest, "image-size-exceeded");
				actionResponse.setRenderParameter("jspPage",
						LibraryConstants.PAGE_UPLOAD);
			}
		}

		File sampleChapterFile = uploadRequest.getFile("sampleChapter");
		if (sampleChapterFile.getTotalSpace() > 0) {
			serviceContext.setAttribute("SAMPLE_CHAPTER", sampleChapterFile);
			String fileName = uploadRequest.getFileName("sampleChapter");
			serviceContext.setAttribute("FILE_NAME", fileName);
		}

		long bookId = ParamUtil.getLong(uploadRequest, "bookId");
		LMSBookLocalServiceUtil.attachFiles(bookId, serviceContext);

		if (SessionErrors.isEmpty(actionRequest)) {
			// redirecting to original list page
			actionResponse.sendRedirect(ParamUtil.getString(uploadRequest,
					"redirectURL"));
		}
	}

	public void serveResource(ResourceRequest resourceRequest,
			ResourceResponse resourceResponse) throws IOException,
			PortletException {

		String cmd = ParamUtil.getString(resourceRequest, Constants.CMD);

		if (cmd.equalsIgnoreCase("serveImage")) {
			long imageId = ParamUtil.getLong(resourceRequest, "imageId");

			Image image = null;
			try {
				image = ImageLocalServiceUtil.fetchImage(imageId);
			} catch (SystemException e) {
				e.printStackTrace();
			}

			if (Validator.isNotNull(image)) {
				byte[] bytes = image.getTextObj();

				HttpServletResponse response = PortalUtil
						.getHttpServletResponse(resourceResponse);

				bytes = LMSUtil.getScaledImage(image, 20);

				response.setContentType(image.getType());
				ServletResponseUtil.write(response, bytes);
			}
		}

		if (cmd.equalsIgnoreCase("serveFile")) {
			String fileName = ParamUtil.getString(resourceRequest, "fileName");

			long companyId = PortalUtil.getCompanyId(resourceRequest);
			long repositoryId = CompanyConstants.SYSTEM;
			String filePath = "Sample_Chapters" + StringPool.SLASH + fileName;

			File file = null;
			try {
				file = DLStoreUtil.getFile(companyId, repositoryId, filePath);
			} catch (PortalException e) {
				e.printStackTrace();
			} catch (SystemException e) {
				e.printStackTrace();
			}

			if (Validator.isNotNull(file)) {
				byte[] bytes = FileUtil.getBytes(file);
				String contentType = new MimetypesFileTypeMap()
						.getContentType(file);

				// preparing the response
				HttpServletResponse response = PortalUtil
						.getHttpServletResponse(resourceResponse);
				response.setContentType(contentType);
				response.setHeader("Content-Disposition",
						"attachment; filename= " + fileName);
				ServletResponseUtil.write(response, bytes);
			}
		}

		if (cmd.equalsIgnoreCase(Constants.SUBSCRIBE)
				|| cmd.equalsIgnoreCase(Constants.UNSUBSCRIBE)) {

			LMSUtil.applySubscription(resourceRequest, cmd);

			// setting the response in JSON format
			JSONObject jsonObj = JSONFactoryUtil.createJSONObject();
			jsonObj.put("subscribed", cmd.equalsIgnoreCase(Constants.SUBSCRIBE));
			HttpServletResponse response = PortalUtil
					.getHttpServletResponse(resourceResponse);
			PrintWriter pw = response.getWriter();
			pw.write(jsonObj.toString());
			pw.close();
		}
	}

	public void discussOnThisBook(ActionRequest actionRequest,
			ActionResponse actionResponse) throws IOException, PortletException {

		String className = "com.liferay.portlet.messageboards.action."
				+ "EditDiscussionAction";
		String methodName = "processAction";

		String[] parameterTypeNames = {
				"org.apache.struts.action.ActionMapping",
				"org.apache.struts.action.ActionForm",
				PortletConfig.class.getName(), ActionRequest.class.getName(),
				ActionResponse.class.getName() };

		Object[] arguments = { null, null, getPortletConfig(), actionRequest,
				actionResponse };

		try {
			PortalClassInvoker.invoke(true, className, methodName,
					parameterTypeNames, arguments);
		} catch (Exception e) {
			e.printStackTrace();
		}		
	}
	
	public void doView(RenderRequest renderRequest,
			RenderResponse renderResponse) throws IOException, PortletException {
		
		ThemeDisplay themeDisplay = (ThemeDisplay) 
				renderRequest.getAttribute(WebKeys.THEME_DISPLAY);
		
		long companyId = themeDisplay.getCompanyId();
		long groupId = themeDisplay.getScopeGroupId();
		
		List<LMSBook> books = LMSBookLocalServiceUtil.advancedSearch(
			companyId, groupId, "java", "veena", true);
		
		System.out.println(books);
		
		super.doView(renderRequest, renderResponse);
	}
}