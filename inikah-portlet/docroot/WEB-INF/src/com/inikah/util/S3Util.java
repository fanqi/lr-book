package com.inikah.util;

import java.io.ByteArrayInputStream;

import org.jets3t.service.S3Service;
import org.jets3t.service.S3ServiceException;
import org.jets3t.service.ServiceException;
import org.jets3t.service.acl.AccessControlList;
import org.jets3t.service.impl.rest.httpclient.RestS3Service;
import org.jets3t.service.model.S3Bucket;
import org.jets3t.service.model.S3Object;
import org.jets3t.service.security.AWSCredentials;

public class S3Util {

	static S3Service client = null;
	
	static final String AWS_ACCESS_KEY = AppConfig.get(IConstants.CFG_AWS_ACCESS_KEY);
	static final String AWS_SECRET_KEY = AppConfig.get(IConstants.CFG_AWS_SECRET_KEY);

	public static void uploadToS3(String bucketName, byte[] bytes, String contentType,
			String storagePath) {
		
		S3Bucket bucket = new S3Bucket(bucketName);
		
		S3Object object = new S3Object();
		object.setBucketName(bucketName);
		object.setDataInputStream(new ByteArrayInputStream(bytes));
		object.setKey(storagePath);
		object.setContentType("image/" + contentType);
		object.setAcl(AccessControlList.REST_CANNED_PUBLIC_READ);

		try {
			object = client.putObject(bucket, object);
		} catch (S3ServiceException e) {
			e.printStackTrace();
		}
	}
	
	public static void removeFromS3(String bucketName, String objectKey) {
		try {
			client.deleteObject(bucketName, objectKey);
		} catch (ServiceException e) {
			e.printStackTrace();
		}
	}

	static {
		AWSCredentials awsCredentials = new AWSCredentials(AWS_ACCESS_KEY, AWS_SECRET_KEY);
		try {
			client = new RestS3Service(awsCredentials);
		} catch (S3ServiceException e) {
			e.printStackTrace();
		}
	}
}