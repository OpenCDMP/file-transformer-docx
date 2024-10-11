package org.opencdmp.filetransformer.docx.service.storage;

public interface FileStorageService {
	String storeFile(byte[] data);

	byte[] readFile(String fileRef);
}
