import { parseStringPromise } from 'xml2js';

interface GroupMember {
    uuid: string;
    location: string;
    name?: string;
}

interface ZoneGroup {
    coordinator: string;
    name: string;
    members: string[];
}

/**
 * Sonos UPnP/SOAP API Client that handles stereo pairs and zone groups
 */
export class SonosController {
    private host: string | null = null;
    private port: number = 1400;
    private coordinatorHost: string | null = null;
    private zoneGroupState: unknown = null;
    private zoneGroupStateTimestamp: number = 0;
    private singleSpeakerMode: boolean = false;

    connect(host: string, port: number = 1400, singleSpeakerMode: boolean = false): void {
        this.host = host;
        this.port = port;
        this.singleSpeakerMode = singleSpeakerMode;
        this.coordinatorHost = null;
        this.zoneGroupState = null;
        this.zoneGroupStateTimestamp = 0;
    }

    isConnected(): boolean {
        return !!this.host;
    }

    private extractHostFromLocation(location?: string): string | null {
        const match = location?.match(/http:\/\/([^:]+):/);
        return match ? match[1] : null;
    }

    /**
     * Get the group coordinator for this speaker.
     * For stereo pairs, this finds the coordinator that controls both speakers.
     */
    async getGroupCoordinator(): Promise<string> {
        if (this.coordinatorHost) {
            return this.coordinatorHost;
        }

        this.zoneGroupState = null;

        const groups = await this.getAvailableGroups();
        const targetGroup = groups.find(group =>
            group.members.some(member => member?.includes(this.host!))
        );

        if (targetGroup?.coordinator) {
            this.coordinatorHost = targetGroup.coordinator;
            return this.coordinatorHost;
        }

        // If we couldn't find coordinator or speaker is standalone, use the original host
        this.coordinatorHost = this.host!;
        return this.host!;
    }

    /**
     * Get the zone group state from Sonos
     */
    private async getZoneGroupState(): Promise<Record<string, unknown>> {
        // Cache zone group state for 30 seconds to prevent excessive API calls
        const now = Date.now();
        if (this.zoneGroupState && this.zoneGroupStateTimestamp && (now - this.zoneGroupStateTimestamp < 30000)) {
            return this.zoneGroupState as Record<string, unknown>;
        }

        const result = await this.executeAction('ZoneGroupTopology', 'GetZoneGroupState', {});
        const zoneGroupState = result.ZoneGroupState as string;
        const parsed = await parseStringPromise(zoneGroupState, {
            explicitArray: true,
            mergeAttrs: false
        });
        this.zoneGroupState = parsed;
        this.zoneGroupStateTimestamp = now;
        return parsed;
    }

    private collectGroupMembers(node: unknown, members: GroupMember[]): void {
        if (!node) {
            return;
        }

        if (Array.isArray(node)) {
            node.forEach(item => this.collectGroupMembers(item, members));
            return;
        }

        if (typeof node === 'object' && node !== null) {
            const nodeObj = node as Record<string, unknown>;
            const attrs = nodeObj.$ as Record<string, string> | undefined;
            if (attrs?.UUID && attrs?.Location) {
                members.push({
                    uuid: attrs.UUID,
                    location: attrs.Location,
                    name: attrs.ZoneName || attrs.ZoneGroupName
                });
            }

            Object.values(nodeObj).forEach(value => this.collectGroupMembers(value, members));
        }
    }

    /**
     * Get all available zone groups (including stereo pairs)
     */
    async getAvailableGroups(): Promise<ZoneGroup[]> {
        const state = await this.getZoneGroupState();
        const stateObj = state as { ZoneGroupState?: { ZoneGroups?: Array<{ ZoneGroup?: unknown }> } };
        const groups = stateObj.ZoneGroupState?.ZoneGroups?.[0]?.ZoneGroup || [];
        const groupList = Array.isArray(groups) ? groups : [groups];

        return groupList.map((group) => {
            const members: GroupMember[] = [];
            this.collectGroupMembers(group, members);

            const uniqueMembers = members.filter((member, index, list) =>
                list.findIndex((item) => item.uuid === member.uuid) === index
            );

            const groupObj = group as { $?: { Coordinator?: string; ZoneGroupName?: string } };
            const coordinatorUuid = groupObj.$?.Coordinator;
            const coordinatorMember = uniqueMembers.find(member => member.uuid === coordinatorUuid)
                || uniqueMembers[0];
            const coordinator = this.extractHostFromLocation(coordinatorMember?.location)
                || this.host!;
            const name = groupObj.$?.ZoneGroupName || coordinatorMember?.name || coordinator;
            const memberHosts = uniqueMembers
                .map(member => this.extractHostFromLocation(member.location))
                .filter((host): host is string => host !== null);

            return {
                coordinator,
                name,
                members: [...new Set(memberHosts)]
            };
        });
    }

    /**
     * Get all members of the current speaker's group (for stereo pairs, returns both speakers)
     */
    async getGroupMembers(): Promise<string[]> {
        const groups = await this.getAvailableGroups();
        const target = groups.find(group =>
            group.members.some(member => member?.includes(this.host!))
        );

        return target?.members || [];
    }

    /**
     * Get volume from the group coordinator
     */
    async getVolume(): Promise<number> {
        const coordinator = await this.getGroupCoordinator();
        const result = await this.executeActionOnHost(coordinator, 'RenderingControl', 'GetVolume', { Channel: 'Master' });
        return parseInt(result.CurrentVolume as string, 10);
    }

    /**
     * Set volume. In single speaker mode, only affects this speaker.
     * In group mode, affects ALL members of the group (handles stereo pairs correctly).
     */
    async setVolume(volume: number): Promise<void> {
        // In single speaker mode, only control this speaker
        if (this.singleSpeakerMode) {
            await this.executeAction('RenderingControl', 'SetVolume', {
                Channel: 'Master',
                DesiredVolume: volume
            });
            return;
        }

        const members = await this.getGroupMembers();
        if (!members.length) {
            await this.executeAction('RenderingControl', 'SetVolume', {
                Channel: 'Master',
                DesiredVolume: volume
            });
            return;
        }

        await Promise.all(members.map(member =>
            this.executeActionOnHost(member, 'RenderingControl', 'SetVolume', {
                Channel: 'Master',
                DesiredVolume: volume
            })
        ));
    }

    /**
     * Get mute state
     */
    async getMuted(): Promise<boolean> {
        const result = await this.executeAction('RenderingControl', 'GetMute', { Channel: 'Master' });
        return result.CurrentMute === '1' || result.CurrentMute === true;
    }

    /**
     * Set mute state. In single speaker mode, only affects this speaker.
     * In group mode, affects ALL members of the group (handles stereo pairs correctly).
     */
    async setMuted(muted: boolean): Promise<void> {
        // In single speaker mode, only control this speaker
        if (this.singleSpeakerMode) {
            await this.executeAction('RenderingControl', 'SetMute', {
                Channel: 'Master',
                DesiredMute: muted ? '1' : '0'
            });
            return;
        }

        const members = await this.getGroupMembers();
        if (!members.length) {
            await this.executeAction('RenderingControl', 'SetMute', {
                Channel: 'Master',
                DesiredMute: muted ? '1' : '0'
            });
            return;
        }

        await Promise.all(members.map(member =>
            this.executeActionOnHost(member, 'RenderingControl', 'SetMute', {
                Channel: 'Master',
                DesiredMute: muted ? '1' : '0'
            })
        ));
    }

    /**
     * Execute a SOAP action on a specific host
     */
    private async executeActionOnHost(
        host: string,
        service: string,
        action: string,
        params: Record<string, unknown>
    ): Promise<Record<string, unknown>> {
        params.InstanceID = params.InstanceID ?? 0;

        const serviceMap: Record<string, string> = {
            'AVTransport': 'MediaRenderer/AVTransport',
            'RenderingControl': 'MediaRenderer/RenderingControl',
            'ZoneGroupTopology': 'ZoneGroupTopology',
            'ContentDirectory': 'MediaServer/ContentDirectory'
        };

        const baseUrl = serviceMap[service] || service;
        const url = `http://${host}:${this.port}/${baseUrl}/Control`;
        const soapAction = `"urn:schemas-upnp-org:service:${service}:1#${action}"`;

        const xmlParams = Object.keys(params)
            .map(key => `<${key}>${this.escapeXml(params[key])}</${key}>`)
            .join('');

        const request = `<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
    <s:Body>
        <u:${action} xmlns:u="urn:schemas-upnp-org:service:${service}:1">
            ${xmlParams}
        </u:${action}>
    </s:Body>
</s:Envelope>`;

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'SOAPAction': soapAction,
                'Content-Type': 'text/xml; charset=utf-8'
            },
            body: request
        });

        const responseText = await response.text();

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${responseText}`);
        }

        const result = await parseStringPromise(responseText, {
            explicitArray: false,
            tagNameProcessors: [(name: string) => name.replace(/^.*:/, '')] // Remove namespace prefixes
        });

        const responseBody = result.Envelope?.Body?.[`${action}Response`];
        return responseBody || {};
    }

    /**
     * Execute a SOAP action on the connected host
     */
    private async executeAction(
        service: string,
        action: string,
        params: Record<string, unknown>
    ): Promise<Record<string, unknown>> {
        if (!this.isConnected()) {
            throw new Error('Not connected to Sonos');
        }

        return this.executeActionOnHost(this.host!, service, action, params);
    }

    private escapeXml(value: unknown): string {
        if (value == null) return '';
        return String(value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }
}
