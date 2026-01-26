import { action, DialAction, DialRotateEvent, SingletonAction, WillAppearEvent, DialUpEvent, TouchTapEvent, DidReceiveSettingsEvent, WillDisappearEvent } from '@elgato/streamdeck';
import streamDeck from '@elgato/streamdeck';
import { SonosController } from '../sonos-controller';

/**
 * Sonos Volume Dial action that controls a Sonos speaker's volume.
 * Supports stereo pairs and zone groups by controlling all group members.
 */
@action({ UUID: 'com.moonbeamalpha.sonos-volume-dial.volume' })
export class SonosVolumeDial extends SingletonAction {
	// Constants
	private static readonly POLLING_INTERVAL_MS = 3000;
	private static readonly VOLUME_CHANGE_DEBOUNCE_MS = 500;

	private logger = streamDeck.logger.createScope('SonosVolumeDial');
	private actionStates = new Map<string, ActionState>();

	private ensureActionState(dialAction: DialAction<SonosVolumeDialSettings>): ActionState {
		const actionId = dialAction.id;
		const existingState = this.actionStates.get(actionId);
		if (existingState) {
			return existingState;
		}

		const newState: ActionState = {
			sonos: null,
			lastKnownVolume: 50,
			isMuted: false,
			pollInterval: null,
			pollTimeoutId: null,
			currentSettings: null,
			volumeChangeTimeout: null,
			isRotating: false
		};
		this.actionStates.set(actionId, newState);
		return newState;
	}

	private stopPollingForState(state: ActionState) {
		if (state.pollInterval) {
			this.logger.debug('Stopping polling');
			state.pollInterval.active = false;
			state.pollInterval = null;
		}
		if (state.pollTimeoutId) {
			clearTimeout(state.pollTimeoutId);
			state.pollTimeoutId = null;
		}
	}

	private cleanupAction(actionId: string) {
		const state = this.actionStates.get(actionId);
		if (!state) {
			return;
		}
		if (state.volumeChangeTimeout) {
			clearTimeout(state.volumeChangeTimeout);
			state.volumeChangeTimeout = null;
		}
		this.stopPollingForState(state);
		state.sonos = null;
		this.actionStates.delete(actionId);
	}

	/**
	 * Start polling for speaker state
	 */
	private startPolling(dialAction: DialAction<SonosVolumeDialSettings>, state: ActionState) {
		// Create a scoped logger for polling
		const logger = this.logger.createScope('Polling');

		// Only start polling if there isn't already an active poll
		if (state.pollInterval?.active) {
			logger.debug('Polling already active, skipping');
			return;
		}

		// Clear any existing poll interval just in case, but preserve state
		if (state.pollInterval) {
			state.pollInterval.active = false;
			state.pollInterval = null;
		}

		// Clear any existing timeout
		if (state.pollTimeoutId) {
			clearTimeout(state.pollTimeoutId);
			state.pollTimeoutId = null;
		}

		// Verify we have necessary state to start polling
		if (!state.currentSettings) {
			logger.debug('Missing required state, cannot start polling');
			return;
		}

		// Start polling using self-scheduling
		state.pollInterval = { active: true };
		logger.debug('Starting polling');
		this.pollWithDelay(dialAction, state, logger);
	}

	/**
	 * Show an alert to the user
	 */
	private showAlert(action: DialAction<SonosVolumeDialSettings>, message: string) {
		action.showAlert();
		this.logger.error(message);
	}

	/**
	 * Self-scheduling poll function that maintains consistent spacing
	 */
	private async pollWithDelay(
		dialAction: DialAction<SonosVolumeDialSettings>,
		state: ActionState,
		logger: ReturnType<typeof streamDeck.logger.createScope>
	) {
		// Ensure we're not running multiple polling cycles
		if (!state.pollInterval?.active) {
			return;
		}

		try {
			if (!state.currentSettings) {
				logger.debug('No current action or settings, stopping polling');
				this.stopPollingForState(state);
				return;
			}

			try {
				// If we don't have a connection, try to reconnect
				if (!state.sonos || !state.sonos.isConnected()) {
					if (state.currentSettings.speakerIp) {
						logger.info('Reconnecting to speaker:', state.currentSettings.speakerIp);
						state.sonos = new SonosController();
						state.sonos.connect(state.currentSettings.speakerIp, 1400, state.currentSettings.singleSpeakerMode ?? false);
					} else {
						logger.debug('No speaker IP, stopping polling');
						this.stopPollingForState(state);
						return;
					}
				}

				// Get current volume and mute state
				const [volume, isMuted] = await Promise.all([
					state.sonos.getVolume(),
					state.sonos.getMuted()
				]);

				// Only update if values have changed and we're not actively rotating
				if ((volume !== state.lastKnownVolume || isMuted !== state.isMuted) && !state.isRotating) {
					logger.debug('Speaker state changed externally - volume:', volume, 'muted:', isMuted);
					state.lastKnownVolume = volume;
					state.isMuted = isMuted;

					// Update UI to reflect current state
					dialAction.setFeedback({
						value: {
							value: volume,
							opacity: isMuted ? 0.5 : 1.0
						},
						indicator: {
							value: volume,
							opacity: isMuted ? 0.5 : 1.0
						}
					});
					state.currentSettings = { ...state.currentSettings, value: volume };
					dialAction.setSettings(state.currentSettings);
				}
			} catch (error) {
				logger.error('Failed to poll speaker state:', {
					error: error instanceof Error ? error.message : String(error),
					stack: error instanceof Error ? error.stack : undefined
				});
				// Don't stop polling on error, just clear the connection so we'll try to reconnect next time
				state.sonos = null;
			}
		} finally {
			// Schedule next poll only if polling is still active
			if (state.pollInterval?.active) {
				if (state.pollTimeoutId) {
					clearTimeout(state.pollTimeoutId);
				}
				state.pollTimeoutId = setTimeout(() => {
					state.pollTimeoutId = null;
					if (state.pollInterval?.active) {
						this.pollWithDelay(dialAction, state, logger);
					}
				}, SonosVolumeDial.POLLING_INTERVAL_MS);
			}
		}
	}

	/**
	 * Stop polling for speaker state
	 */
	// Per-action cleanup happens in onWillDisappear

	/**
	 * Sets the initial value when the action appears on Stream Deck.
	 */
	override async onWillAppear(ev: WillAppearEvent<SonosVolumeDialSettings>): Promise<void> {
		// Create a scoped logger for this specific instance
		const logger = this.logger.createScope('WillAppear');
		
		try {
			// Verify that the action is a dial so we can call setFeedback.
			if (!ev.action.isDial()) return;

			const dialAction = ev.action as DialAction<SonosVolumeDialSettings>;
			const { speakerIp, value = 50, volumeStep = 5, singleSpeakerMode = false } = ev.payload.settings;
			const state = this.ensureActionState(dialAction);

			// Store current settings for this action
			state.currentSettings = ev.payload.settings;

			// Initialize display with current or default value
			dialAction.setFeedback({
				value: {
					value,
					opacity: state.isMuted ? 0.5 : 1.0
				},
				indicator: {
					value,
					opacity: state.isMuted ? 0.5 : 1.0
				}
			});

			// If we have a speaker IP, initialize the connection and update volume
			if (speakerIp) {
				logger.info('Connecting to Sonos speaker:', speakerIp, singleSpeakerMode ? '(single speaker mode)' : '(group mode)');
				state.sonos = new SonosController();
				state.sonos.connect(speakerIp, 1400, singleSpeakerMode);
				
				try {
					// Get current volume and mute state
					const [volume, isMuted] = await Promise.all([
						state.sonos.getVolume(),
						state.sonos.getMuted()
					]);
					
					state.lastKnownVolume = volume;
					state.isMuted = isMuted;
					
					// Update UI with current state
					dialAction.setFeedback({ 
						value: {
							value: volume,
							opacity: isMuted ? 0.5 : 1.0
						},
						indicator: { 
							value: volume,
							opacity: isMuted ? 0.5 : 1.0
						}
					});

					// Send settings back to Property Inspector with current volume
					state.currentSettings = { speakerIp, volumeStep, value: volume };
					dialAction.setSettings(state.currentSettings);

					// Start polling for updates only after we've successfully connected and initialized
					this.startPolling(dialAction, state);
				} catch (error) {
					logger.error('Failed to connect to speaker:', {
						error: error instanceof Error ? error.message : String(error),
						stack: error instanceof Error ? error.stack : undefined
					});
					state.sonos = null;
					this.showAlert(dialAction, 'Failed to connect to speaker');
					// Even if connection fails, ensure settings are synced
					state.currentSettings = { speakerIp, volumeStep, value };
					dialAction.setSettings(state.currentSettings);
				}
			} else {
				logger.warn('No speaker IP configured');
				state.lastKnownVolume = value;
				// Ensure settings are synced even when no IP is configured
				state.currentSettings = { volumeStep, value };
				dialAction.setSettings(state.currentSettings);
			}
		} catch (error) {
			logger.error('Error in onWillAppear:', {
				error: error instanceof Error ? error.message : String(error),
				stack: error instanceof Error ? error.stack : undefined
			});
		}
	}

	/**
	 * Update the value based on the dial rotation.
	 */
	override async onDialRotate(ev: DialRotateEvent<SonosVolumeDialSettings>): Promise<void> {
		// Create a scoped logger for this specific rotation event
		const logger = this.logger.createScope('DialRotate');
		const dialAction = ev.action as DialAction<SonosVolumeDialSettings>;
		
		try {
			const state = this.ensureActionState(dialAction);
			const { speakerIp, value = state.lastKnownVolume, volumeStep = 5 } = ev.payload.settings;

			// Mark that we're actively rotating
			state.isRotating = true;

			// Update stored settings
			state.currentSettings = ev.payload.settings;
			
			const { ticks } = ev.payload;

			// Calculate new value using the volumeStep setting
			const newValue = Math.max(0, Math.min(100, value + (ticks * volumeStep)));

			// Update UI immediately for responsiveness
			dialAction.setFeedback({ 
				value: {
					value: newValue,
					opacity: state.isMuted ? 0.5 : 1.0
				},
				indicator: { 
					value: newValue,
					opacity: state.isMuted ? 0.5 : 1.0
				}
			});
			state.currentSettings = { ...state.currentSettings, value: newValue };
			dialAction.setSettings(state.currentSettings);
			state.lastKnownVolume = newValue;

			// Clear any pending volume change
			if (state.volumeChangeTimeout) {
				clearTimeout(state.volumeChangeTimeout);
				state.volumeChangeTimeout = null;
			}

			// Handle Sonos operations in the background after debounce
			if (speakerIp) {
				const actionId = dialAction.id;
				state.volumeChangeTimeout = setTimeout(async () => {
					const currentState = this.actionStates.get(actionId);
					if (!currentState) {
						return;
					}
					try {
						// Initialize connection if needed
						if (!currentState.sonos || !currentState.sonos.isConnected()) {
							logger.info('Reconnecting to speaker:', speakerIp);
							currentState.sonos = new SonosController();
							currentState.sonos.connect(speakerIp, 1400, currentState.currentSettings?.singleSpeakerMode ?? false);
						}

						// If speaker is muted, unmute it first
						if (currentState.isMuted) {
							await currentState.sonos.setMuted(false);
							currentState.isMuted = false;
						}

						// Set the volume without waiting for verification
						await currentState.sonos.setVolume(newValue);
						logger.debug('Volume successfully set to:', newValue);
					} catch (error) {
						logger.error('Failed to update volume:', {
							error: error instanceof Error ? error.message : String(error),
							stack: error instanceof Error ? error.stack : undefined,
							targetVolume: newValue
						});
						currentState.sonos = null;
						this.showAlert(dialAction, 'Failed to update volume');
					} finally {
						// Clear rotating flag and restart polling only after the last debounced update
						currentState.isRotating = false;
						this.startPolling(dialAction, currentState);
					}
				}, SonosVolumeDial.VOLUME_CHANGE_DEBOUNCE_MS);
			} else {
				logger.warn('No speaker IP configured');
				this.showAlert(dialAction, 'No speaker IP configured');
				state.isRotating = false;
			}
		} catch (error) {
			logger.error('Error in onDialRotate:', {
				error: error instanceof Error ? error.message : String(error),
				stack: error instanceof Error ? error.stack : undefined
			});
			const state = this.ensureActionState(dialAction);
			state.isRotating = false;
		}
	}

	/**
	 * Toggle mute state when the dial is pressed.
	 */
	override async onDialUp(ev: DialUpEvent<SonosVolumeDialSettings>): Promise<void> {
		// Create a scoped logger for this specific press event
		const logger = this.logger.createScope('DialUp');
		
		try {
			const dialAction = ev.action as DialAction<SonosVolumeDialSettings>;
			const state = this.ensureActionState(dialAction);
			const { speakerIp } = ev.payload.settings;

			// Update UI immediately with optimistic state
			const newMutedState = !state.isMuted;
			state.isMuted = newMutedState;
			dialAction.setFeedback({ 
				value: {
					value: state.lastKnownVolume,
					opacity: newMutedState ? 0.5 : 1.0,
				},
				indicator: { 
					value: state.lastKnownVolume,
					opacity: newMutedState ? 0.5 : 1.0
				}
			});

			// Handle Sonos operations in the background
			if (speakerIp) {
				const singleSpeakerMode = state.currentSettings?.singleSpeakerMode ?? false;
				Promise.resolve().then(async () => {
					try {
						// Initialize connection if needed
						if (!state.sonos || !state.sonos.isConnected()) {
							logger.info('Reconnecting to speaker:', speakerIp);
							state.sonos = new SonosController();
							state.sonos.connect(speakerIp, 1400, singleSpeakerMode);
							// Restart polling if it was stopped
							if (!state.pollInterval) {
								this.startPolling(dialAction, state);
							}
						}

						// Set mute state without waiting for verification
						// Let the polling cycle handle any discrepancies
						await state.sonos.setMuted(newMutedState);
					} catch (error) {
						logger.error('Failed to toggle mute:', {
							error: error instanceof Error ? error.message : String(error),
							stack: error instanceof Error ? error.stack : undefined
						});
						state.sonos = null;
						this.showAlert(dialAction, 'Failed to toggle mute');
						// Keep optimistic update UI state, let polling sync actual state
					}
				});
			} else {
				logger.warn('No speaker IP configured');
				this.showAlert(dialAction, 'No speaker IP configured');
			}
		} catch (error) {
			logger.error('Error in onDialUp:', {
				error: error instanceof Error ? error.message : String(error),
				stack: error instanceof Error ? error.stack : undefined
			});
		}
	}

	/**
	 * Toggle mute state when the dial face is tapped.
	 */
	override async onTouchTap(ev: TouchTapEvent<SonosVolumeDialSettings>): Promise<void> {
		// Create a scoped logger for this specific tap event
		const logger = this.logger.createScope('TouchTap');
		
		try {
			const dialAction = ev.action as DialAction<SonosVolumeDialSettings>;
			const state = this.ensureActionState(dialAction);
			const { speakerIp } = ev.payload.settings;

			// Update UI immediately with optimistic state
			const newMutedState = !state.isMuted;
			state.isMuted = newMutedState;
			dialAction.setFeedback({ 
				value: {
					value: state.lastKnownVolume,
					opacity: newMutedState ? 0.5 : 1.0,
				},
				indicator: { 
					value: state.lastKnownVolume,
					opacity: newMutedState ? 0.5 : 1.0
				}
			});

			// Handle Sonos operations in the background
			if (speakerIp) {
				const singleSpeakerMode = state.currentSettings?.singleSpeakerMode ?? false;
				Promise.resolve().then(async () => {
					try {
						// Initialize connection if needed
						if (!state.sonos || !state.sonos.isConnected()) {
							logger.info('Reconnecting to speaker:', speakerIp);
							state.sonos = new SonosController();
							state.sonos.connect(speakerIp, 1400, singleSpeakerMode);
							// Restart polling if it was stopped
							if (!state.pollInterval) {
								this.startPolling(dialAction, state);
							}
						}

						// Set mute state without waiting for verification
						// Let the polling cycle handle any discrepancies
						await state.sonos.setMuted(newMutedState);
					} catch (error) {
						logger.error('Failed to toggle mute:', {
							error: error instanceof Error ? error.message : String(error),
							stack: error instanceof Error ? error.stack : undefined
						});
						state.sonos = null;
						this.showAlert(dialAction, 'Failed to toggle mute');
						// Keep optimistic update UI state, let polling sync actual state
					}
				});
			} else {
				logger.warn('No speaker IP configured');
				this.showAlert(dialAction, 'No speaker IP configured');
			}
		} catch (error) {
			logger.error('Error in onTouchTap:', {
				error: error instanceof Error ? error.message : String(error),
				stack: error instanceof Error ? error.stack : undefined
			});
		}
	}

	/**
	 * Handle settings updates
	 */
	override async onDidReceiveSettings(ev: DidReceiveSettingsEvent<SonosVolumeDialSettings>): Promise<void> {
		const logger = this.logger.createScope('DidReceiveSettings');
		
		try {
			if (!ev.action.isDial()) return;

			const dialAction = ev.action as DialAction<SonosVolumeDialSettings>;
			const state = this.ensureActionState(dialAction);
			const previousSettings = state.currentSettings;
			const { speakerIp, value = state.lastKnownVolume, volumeStep = 5, singleSpeakerMode = false } = ev.payload.settings;

			// Store current settings
			state.currentSettings = ev.payload.settings;

			// If speaker IP or single speaker mode changed, we need to reconnect
			if (speakerIp !== previousSettings?.speakerIp || singleSpeakerMode !== previousSettings?.singleSpeakerMode) {
				// Clear existing connection
				state.sonos = null;
				this.stopPollingForState(state);

				if (speakerIp) {
					logger.info('Connecting to speaker:', speakerIp, singleSpeakerMode ? '(single speaker mode)' : '(group mode)');
					state.sonos = new SonosController();
					state.sonos.connect(speakerIp, 1400, singleSpeakerMode);
					
					try {
						// Get current volume and mute state
						const [volume, isMuted] = await Promise.all([
							state.sonos.getVolume(),
							state.sonos.getMuted()
						]);
						
						state.lastKnownVolume = volume;
						state.isMuted = isMuted;
						
						// Update UI with current state
						dialAction.setFeedback({ 
							value: {
								value: volume,
								opacity: isMuted ? 0.5 : 1.0
							},
							indicator: { 
								value: volume,
								opacity: isMuted ? 0.5 : 1.0
							}
						});
						state.currentSettings = { ...ev.payload.settings, value: volume };
						dialAction.setSettings(state.currentSettings);

						// Start polling for updates
						this.startPolling(dialAction, state);
					} catch (error) {
						logger.error('Failed to connect to new speaker:', {
							error: error instanceof Error ? error.message : String(error),
							stack: error instanceof Error ? error.stack : undefined
						});
						state.sonos = null;
						this.showAlert(dialAction, 'Failed to connect to speaker');
					}
				} else {
					state.lastKnownVolume = value;
				}
			}
		} catch (error) {
			logger.error('Error in onDidReceiveSettings:', {
				error: error instanceof Error ? error.message : String(error),
				stack: error instanceof Error ? error.stack : undefined
			});
		}
	}

	/**
	 * Clean up when the action is removed
	 */
	override onWillDisappear(ev: WillDisappearEvent<SonosVolumeDialSettings>): void {
		this.cleanupAction(ev.action.id);
	}
}

/**
 * Settings for {@link SonosVolumeDial}.
 */
type SonosVolumeDialSettings = {
	value: number;
	speakerIp?: string;
	volumeStep: number;
	singleSpeakerMode?: boolean;
};

type ActionState = {
	sonos: SonosController | null;
	lastKnownVolume: number;
	isMuted: boolean;
	pollInterval: { active: boolean } | null;
	pollTimeoutId: NodeJS.Timeout | null;
	currentSettings: SonosVolumeDialSettings | null;
	volumeChangeTimeout: NodeJS.Timeout | null;
	isRotating: boolean;
};
