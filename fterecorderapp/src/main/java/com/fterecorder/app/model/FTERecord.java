package com.fterecorder.app.model;

import java.io.Serializable;
import java.time.LocalDateTime;


import org.springframework.data.annotation.Id;
import org.springframework.data.annotation.PersistenceConstructor;
import org.springframework.data.annotation.Transient;
import org.springframework.data.mongodb.core.mapping.Document;
import org.springframework.format.annotation.DateTimeFormat;


@Document(collection = "FTERecord")
public class FTERecord implements Serializable {

	@Transient
	public static final String SEQUENCE_NAME = "fterecord_seq";
	@Id
	long id;
	String track;
	float ftesPerRosterCount;
	float defectCount;
	float widgetCount;
	float ftesLoanedCount;
	float ptoCount;
	float ftesBorrowedCount;
	float ftesForPerformanceCount;
	float ftesForExtendedScenarioExecCount;
	float excessFteCount;
	float medsDefAnalysisCount;
	float workableDefectsCount;
	float defectFteCount;
	float ndeDefectAnalysis;

	@DateTimeFormat(iso = DateTimeFormat.ISO.DATE_TIME)
	LocalDateTime createDt;
	@DateTimeFormat(iso = DateTimeFormat.ISO.DATE_TIME)
	LocalDateTime weekStDt;
	@DateTimeFormat(iso = DateTimeFormat.ISO.DATE_TIME)
	LocalDateTime weekEdDt;
	String note;
	Employee employee;

	public FTERecord() {
	}

	@PersistenceConstructor
	public FTERecord(long id, String track, float ftesPerRosterCount, float defectCount, float defectFteCount,
			float widgetCount, float ftesLoanedCount, float ptoCount, float ftesBorrowedCount, float ftesForPerformanceCount,
			float ftesForExtendedScenarioExecCount, float excessFteCount, float workableDefectsCount,
			float medsDefAnalysisCount, @DateTimeFormat(iso = DateTimeFormat.ISO.DATE_TIME) LocalDateTime createDt, @DateTimeFormat(iso = DateTimeFormat.ISO.DATE_TIME) LocalDateTime weekStDt, @DateTimeFormat(iso = DateTimeFormat.ISO.DATE_TIME) LocalDateTime weekEdDt, String note) {

		super();
		this.id = id;
		this.track = track;
		this.ftesPerRosterCount = ftesPerRosterCount;
		this.defectCount = defectCount;
		this.defectFteCount = defectFteCount;
		this.widgetCount = widgetCount;
		this.ftesLoanedCount = ftesLoanedCount;
		this.ptoCount = ptoCount;
		this.ftesBorrowedCount = ftesBorrowedCount;
		this.ftesForPerformanceCount = ftesForPerformanceCount;
		this.ftesForExtendedScenarioExecCount = ftesForExtendedScenarioExecCount;
		this.excessFteCount = excessFteCount;
		this.workableDefectsCount = workableDefectsCount;
		this.medsDefAnalysisCount = medsDefAnalysisCount;
		this.createDt = createDt;
		this.weekStDt = weekStDt;
		this.weekEdDt = weekEdDt;
		this.note = note;
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

	public float getFtesPerRosterCount() {
		return ftesPerRosterCount;
	}

	public void setFtesPerRosterCount(float ftesPerRosterCount) {
		this.ftesPerRosterCount = ftesPerRosterCount;
	}

	public float getDefectCount() {
		return defectCount;
	}

	public void setDefectCount(float defectCount) {
		this.defectCount = defectCount;
	}

	public float getWidgetCount() {
		return widgetCount;
	}

	public void setWidgetCount(float widgetCount) {
		this.widgetCount = widgetCount;
	}

	public float getFtesLoanedCount() {
		return ftesLoanedCount;
	}

	public void setFtesLoanedCount(float ftesLoanedCount) {
		this.ftesLoanedCount = ftesLoanedCount;
	}

	public float getPtoCount() {
		return ptoCount;
	}

	public void setPtoCount(float ptoCount) {
		this.ptoCount = ptoCount;
	}

	public float getFtesBorrowedCount() {
		return ftesBorrowedCount;
	}

	public void setFtesBorrowedCount(float ftesBorrowedCount) {
		this.ftesBorrowedCount = ftesBorrowedCount;
	}

	public float getFtesForPerformanceCount() {
		return ftesForPerformanceCount;
	}

	public void setFtesForPerformanceCount(float ftesForPerformanceCount) {
		this.ftesForPerformanceCount = ftesForPerformanceCount;
	}

	public float getFtesForExtendedScenarioExecCount() {
		return ftesForExtendedScenarioExecCount;
	}

	public void setFtesForExtendedScenarioExecCount(float ftesForExtendedScenarioExecCount) {
		this.ftesForExtendedScenarioExecCount = ftesForExtendedScenarioExecCount;
	}

	public float getExcessFteCount() {
		return excessFteCount;
	}

	public void setExcessFteCount(float excessFteCount) {
		this.excessFteCount = excessFteCount;
	}

	public float getMedsDefAnalysisCount() {
		return medsDefAnalysisCount;
	}

	public void setMedsDefAnalysisCount(float medsDefAnalysisCount) {
		this.medsDefAnalysisCount = medsDefAnalysisCount;
	}

	public float getWorkableDefectsCount() {
		return workableDefectsCount;
	}

	public void setWorkableDefectsCount(float workableDefectsCount) {
		this.workableDefectsCount = workableDefectsCount;
	}

	public float getDefectFteCount() {
		return defectFteCount;
	}

	public void setDefectFteCount(float defectFteCount) {
		this.defectFteCount = defectFteCount;
	}

	public float getNdeDefectAnalysis() {
		return ndeDefectAnalysis;
	}

	public void setNdeDefectAnalysis(float ndeDefectAnalysis) {
		this.ndeDefectAnalysis = ndeDefectAnalysis;
	}

	public LocalDateTime getCreateDt() {
		return createDt;
	}

	public void setCreateDt(LocalDateTime createDt) {
		this.createDt = createDt;
	}

	public LocalDateTime getWeekStDt() {
		return weekStDt;
	}

	public void setWeekStDt(LocalDateTime weekStDt) {
		this.weekStDt = weekStDt;
	}

	public LocalDateTime getWeekEdDt() {
		return weekEdDt;
	}

	public void setWeekEdDt(LocalDateTime weekEdDt) {
		this.weekEdDt = weekEdDt;
	}

	public String getNote() {
		return note;
	}

	public void setNote(String note) {
		this.note = note;
	}

	public Employee getEmployee() {
		return employee;
	}

	public void setEmployee(Employee employee) {
		this.employee = employee;
	}


}
