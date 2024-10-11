package org.opencdmp.filetransformer.docx.service.wordfiletransformer.visibility;

import org.opencdmp.commonmodels.models.description.VisibilityStateModel;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class VisibilityServiceImpl implements VisibilityService {
	private final Map<FieldKey, Boolean> visibility;
	
    public VisibilityServiceImpl(List<VisibilityStateModel> visibilityStates) {
	    this.visibility = new HashMap<>();
	    for (VisibilityStateModel visibilityState : visibilityStates) this.visibility.put(new FieldKey(visibilityState.getFieldId(), visibilityState.getOrdinal()), visibilityState.getVisible());
    }

	@Override
	public boolean isVisible(String id, Integer ordinal) {
		FieldKey fieldKey = new FieldKey(id, ordinal);
		return this.visibility.getOrDefault(fieldKey, false);
	}
}

