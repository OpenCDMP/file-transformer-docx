package org.opencdmp.filetransformer.docx.service.wordfiletransformer;

import org.springframework.boot.context.properties.ConfigurationProperties;

@ConfigurationProperties(prefix = "word-file-transformer")
public class WordFileTransformerServiceProperties {
	private String transformerId;
	private boolean useSharedStorage;
	private String wordPlanTemplate;
	private String pidTemplate;
	private String wordDescriptionTemplate;
	private String organizationReferenceCode;
	private String grantReferenceCode;
	private String funderReferenceCode;
	private String researcherReferenceCode;
	private String licenceReferenceCode;
	private String projectReferenceCode;
	private String datasetReferenceCode;
	private String publicationReferenceCode;

	public String getTransformerId() {
		return transformerId;
	}

	public void setTransformerId(String transformerId) {
		this.transformerId = transformerId;
	}

	public String getOrganizationReferenceCode() {
		return organizationReferenceCode;
	}

	public void setOrganizationReferenceCode(String organizationReferenceCode) {
		this.organizationReferenceCode = organizationReferenceCode;
	}

	public String getGrantReferenceCode() {
		return grantReferenceCode;
	}

	public void setGrantReferenceCode(String grantReferenceCode) {
		this.grantReferenceCode = grantReferenceCode;
	}

	public String getFunderReferenceCode() {
		return funderReferenceCode;
	}

	public void setFunderReferenceCode(String funderReferenceCode) {
		this.funderReferenceCode = funderReferenceCode;
	}

	public String getResearcherReferenceCode() {
		return researcherReferenceCode;
	}

	public void setResearcherReferenceCode(String researcherReferenceCode) {
		this.researcherReferenceCode = researcherReferenceCode;
	}

	public String getLicenceReferenceCode() {
		return licenceReferenceCode;
	}

	public void setLicenceReferenceCode(String licenceReferenceCode) {
		this.licenceReferenceCode = licenceReferenceCode;
	}

	public String getProjectReferenceCode() {
		return projectReferenceCode;
	}

	public void setProjectReferenceCode(String projectReferenceCode) {
		this.projectReferenceCode = projectReferenceCode;
	}

	public String getDatasetReferenceCode() {
		return datasetReferenceCode;
	}

	public void setDatasetReferenceCode(String datasetReferenceCode) {
		this.datasetReferenceCode = datasetReferenceCode;
	}

	public String getPublicationReferenceCode() {
		return publicationReferenceCode;
	}

	public void setPublicationReferenceCode(String publicationReferenceCode) {
		this.publicationReferenceCode = publicationReferenceCode;
	}

	public String getWordPlanTemplate() {
		return wordPlanTemplate;
	}

	public void setWordPlanTemplate(String wordPlanTemplate) {
		this.wordPlanTemplate = wordPlanTemplate;
	}

	public String getPidTemplate() {
		return pidTemplate;
	}

	public void setPidTemplate(String pidTemplate) {
		this.pidTemplate = pidTemplate;
	}

	public String getWordDescriptionTemplate() {
		return wordDescriptionTemplate;
	}

	public void setWordDescriptionTemplate(String wordDescriptionTemplate) {
		this.wordDescriptionTemplate = wordDescriptionTemplate;
	}

	public boolean isUseSharedStorage() {
		return useSharedStorage;
	}

	public void setUseSharedStorage(boolean useSharedStorage) {
		this.useSharedStorage = useSharedStorage;
	}
}
