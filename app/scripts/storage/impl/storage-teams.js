import { StorageOneDrive } from 'storage/impl/storage-onedrive';
import { OneDriveApps } from 'const/cloud-storage-apps';
import { Features } from 'util/features';
import * as microsoftTeams from '@microsoft/teams-js'; // eslint-disable-line import/no-namespace

class StorageTeams extends StorageOneDrive {
    name = 'teams';
    enabled = true;
    uipos = 50;
    // Free icon provided by Icons8
    // https://icons8.com/icon/blLagk1rxZGp/microsoft-teams-2019
    iconSvg = 'teams';

    _getOAuthConfig() {
        let clientId = this.appSettings.teamsClientId;
        if (!clientId) {
            if (Features.isLocal) {
                ({ id: clientId } = OneDriveApps.Local);
            } else {
                ({ id: clientId } = OneDriveApps.Production);
            }
        }
        return {
            url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
            tokenUrl: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            scope: 'files.readwrite sites.readwrite.all offline_access',
            clientId,
            clientSecret: '',
            pkce: true,
            width: 600,
            height: 500
        };
    }

    _openPopup(url, title, width, height, extras) {
        const params = new URLSearchParams(window.location.search);
        // Teams auth is mutually exclusive with the regular auth flow:
        // It only works inside Teams while the regular flow does not work
        // in Teams
        if (params.get('teamsAuth') !== 'true') {
            this.logger.debug('Teams: Calling parent');
            return super._openPopup(url, title, width, height, extras);
        }
        // Save it for later branching in oauth-result
        localStorage.setItem('teamsAuth', 'true');
        this.logger.debug('Teams: Waiting for Teams initialization');

        localStorage.setItem('teams.authUrl', url);

        microsoftTeams.initialize(() => {
            this.logger.debug('Teams: Teams initialized; init auth workflow');

            microsoftTeams.authentication.authenticate({
                url: window.location.origin + '/teams/start-auth.html',
                width,
                height,
                successCallback: (result) => {
                    this.logger.debug('Teams: Auth successful, posting to parent');
                    window.postMessage(result, window.location.origin);
                },
                failureCallback: (reason) => {
                    this.logger.error('Teams: Auth failed, posting to parent');
                    window.postMessage({ storage: 'teams' }, window.location.origin);
                }
            });
        });
        return true;
    }

    /*
     * Looks up the graph URL given a public URL
     * @param path: https://contoso.sharepoint.com/sites/IT/Accounts/IT.kdbx
     * => Callback: /drives/b!6-noXlXoa018L24gfJj5KmM4nm23AzOLrdfgmasnnWMLYmm2vlwMMJEoe20lqLsy/root:/IT.kdbx
     */

    _graphPath(path, callback) {
        const urlMatcher = /https:\/\/([^/]+)(\/[^/]+\/[^/]+)\/([^/]+)(\/.+\.kdbx)/;
        const urlMatches = urlMatcher.exec(path);
        const tenant = urlMatches[1]; // e.g. contoso.sharepoint.com
        const site = urlMatches[2]; // e.g. /sites/IT
        const collection = urlMatches[3]; // e.g. Accounts
        const filePath = urlMatches[4]; // e.g. /IT.kdbx
        if (!tenant || !site || !collection || !filePath) {
            return callback && callback('', 'Invalid URL');
        }
        const url = `https://graph.microsoft.com/v1.0/sites/${tenant}:${site}:/drives?$select=id,webUrl`;
        this._xhr({
            url,
            responseType: 'json',
            success: (response) => {
                const result = response.value.find(
                    (r) => r.webUrl === `https://${tenant}${site}/${collection}`
                );
                if (!result) {
                    return callback && callback('', `Collection "${collection}" not found`);
                }
                const collectionId = result.id;
                const convertedPath = `/drives/${collectionId}/root:${filePath}`;
                callback(convertedPath);
            },
            error: (err) => {
                return callback && callback('', err);
            }
        });
    }

    _withGraphPath(method, callback, path, ...args) {
        this._oauthAuthorize((err) => {
            if (err) {
                return callback && callback(err);
            }
            this._graphPath(path, (convertedPath, err) => {
                if (err) {
                    return callback && callback(err);
                }
                return method.call(this, convertedPath, ...args);
            });
        });
    }

    load(path, opts, callback) {
        return this._withGraphPath(super.load, callback, path, opts, callback);
    }

    stat(path, opts, callback) {
        return this._withGraphPath(super.stat, callback, path, opts, callback);
    }

    save(path, opts, data, callback, rev) {
        return this._withGraphPath(super.save, callback, path, opts, data, callback, rev);
    }
}

export { StorageTeams };
