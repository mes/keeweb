import kdbxweb from 'kdbxweb';
import { Launcher } from 'comp/launcher';
import { AppModel } from 'models/app-model';
import { Alerts } from 'comp/ui/alerts';
import { Locale } from 'util/locale';

const AutoOpener = {
    init() {
        if (Launcher) {
            return;
        }
        const params = new URLSearchParams(window.location.search);
        if (!params.has('storage') || !params.has('path')) {
            return;
        }
        AppModel.instance.openFile(
            {
                id: null,
                name: 'Auto-opened',
                storage: params.get('storage'),
                path: params.get('path'),
                keyFileName: null,
                keyFileData: null,
                keyFilePath: null,
                fileData: null,
                rev: null,
                opts: null,
                chalResp: null,
                password: kdbxweb.ProtectedValue.fromString(params.get('password'))
            },
            (err) => {
                if (err) {
                    Alerts.error({
                        header: Locale.openError,
                        body: Locale.openErrorDescription,
                        pre: err.toString()
                    });
                }
            }
        );
    }
};

export { AutoOpener };
