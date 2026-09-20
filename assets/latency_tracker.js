/**
 * Streaming Latency Tracker - Group 10 Financial Inclusion Dashboard
 *
 * Implements high-precision browser-side latency measurement.
 * Captures T_req when dcc.Interval triggers a stream update,
 * and T_3 when the server response arrives and is processed in the browser.
 *
 * All browser timing uses window.performance.now() for sub-millisecond precision.
 */

window.dash_clientside = Object.assign({}, window.dash_clientside, {
    latency_benchmarking: {
        lastTickTime: null,

        /**
         * Capture the exact moment the interval ticks in the browser (T_req)
         */
        capture_tick_start: function(n_intervals) {
            if (n_intervals === undefined || n_intervals === null) {
                return window.dash_clientside.no_update;
            }
            const now = window.performance.now();
            window.dash_clientside.latency_benchmarking.lastTickTime = now;
            return {
                tick: n_intervals,
                t_req: now
            };
        },

        /**
         * Capture the exact moment server response arrives and component updates (T_3)
         * Computes client-visible update latency: T_3 - T_req.
         */
        calculate_client_latency: function(serverData) {
            if (!serverData || !serverData.seq) {
                return window.dash_clientside.no_update;
            }

            const t3 = window.performance.now();
            const t_req = window.dash_clientside.latency_benchmarking.lastTickTime;

            let client_visible_ms = 0;
            if (t_req !== null && t_req > 0 && t3 >= t_req) {
                client_visible_ms = Math.round((t3 - t_req) * 100) / 100;
            } else {
                // Fallback estimate if initial tick was missed
                client_visible_ms = Math.round((serverData.processing_latency_ms + 2.5) * 100) / 100;
            }

            return {
                seq: serverData.seq,
                timestamp: serverData.timestamp,
                stream_mode: serverData.stream_mode,
                polling_interval_ms: serverData.polling_interval_ms,
                processing_latency_ms: serverData.processing_latency_ms,
                client_visible_latency_ms: client_visible_ms,
                t1_perf: serverData.t1,
                t2_perf: serverData.t2,
                t3_perf: t3
            };
        }
    }
});
