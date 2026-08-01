/**
 * Extension service worker (MV3 module).
 * Loads Level Up (Power Apps / Dynamics) handlers, then SPO Toolkit handlers.
 * Keep Level Up import first so its MessageService registers before SPO listeners.
 */
import "./levelup/background.js";
import "./background.js";
