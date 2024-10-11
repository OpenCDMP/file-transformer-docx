package org.opencdmp.filetransformer.docx.model.interfaces;

/**
 * Created by ikalyvas on 2/27/2018.
 */
public interface ApplierWithValue<A, V, R> {
    R apply(A applier, V value);
}
