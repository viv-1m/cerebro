package com.taxbhai.cerebro.constants;

/**
 * Holding-period thresholds used to classify a realised trade as Short Term
 * vs Long Term. Centralised here so a future tax-rule change (e.g. if gold's
 * threshold ever needs to diverge from equity's again) is a one-line edit
 * instead of a hunt through the parser methods.
 */
public final class CapitalGainsConstants {

    private CapitalGainsConstants() {
        // constants holder, not instantiable
    }

    /**
     * A holding period strictly greater than this many days is Long Term;
     * anything at or below it is Short Term. Applies uniformly to equity,
     * mutual funds, gold, and any other instrument type — there is currently
     * no asset-specific carve-out.
     */
    public static final long LONG_TERM_HOLDING_DAYS_THRESHOLD = 365L;
}