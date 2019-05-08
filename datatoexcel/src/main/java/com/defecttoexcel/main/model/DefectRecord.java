package com.defecttoexcel.main.model;

import org.springframework.data.annotation.Id;
import org.springframework.data.annotation.PersistenceConstructor;
import org.springframework.data.annotation.Transient;
import org.springframework.data.mongodb.core.mapping.Document;

@Document(collection = "DefectRecord")
public class DefectRecord {

	@Transient
	public static final String SEQUENCE_NAME = "defectrecord_seq";
	@Id
	long id;
	String key;

	public String getKey() {
		return key;
	}

	public void setKey(String key) {
		this.key = key;
	}

	String track;
	String summary;
	String applicableToIE;
	String newJiraId;
	String comments;
	String source;

	public String getSource() {
		return source;
	}

	public void setSource(String source) {
		this.source = source;
	}

	public long getId() {
		return id;
	}

	public void setId(long id) {
		this.id = id;
	}

	public String getTrack() {
		return track;
	}

	public void setTrack(String track) {
		this.track = track;
	}

	public String getNewJiraId() {
		return newJiraId;
	}

	public void setNewJiraId(String newJiraId) {
		this.newJiraId = newJiraId;
	}

	public String getComments() {
		return comments;
	}

	public void setComments(String comments) {
		this.comments = comments;
	}

	public String getSummary() {
		return summary;
	}

	public void setSummary(String summary) {
		this.summary = summary;
	}

	public DefectRecord() {
	}

	@PersistenceConstructor
	public DefectRecord(long id, String track, String newJiraId, String comments, String summary, String applicableToIE,
			String key) {
		super();
		this.id = id;
		this.track = track;
		this.newJiraId = newJiraId;
		this.comments = comments;
		this.applicableToIE = applicableToIE;
		this.key = key;
		this.summary = summary;
	}

	public String getApplicableToIE() {
		return applicableToIE;
	}

	public void setApplicableToIE(String applicableToIE) {
		this.applicableToIE = applicableToIE;
	}

	public static String getSequenceName() {
		return SEQUENCE_NAME;
	}

}
