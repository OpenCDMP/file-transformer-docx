package org.opencdmp.filetransformer.docx.service.pid;

import org.opencdmp.filetransformer.docx.model.PidLink;

import java.util.List;

public interface PidService {
	PidLink getPid(String pidType);
	List<PidLink> loadPidLinks();
}
