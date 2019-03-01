(function(hello) {

	hello.init({
		aad: {
			name: 'Azure Active Directory',
			
			oauth: {
				version: 2,
				auth: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
				grant: 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
			},

			// Authorization scopes
			scope: {
				// you can add as many scopes to the mapping as you want here
				profile: 'user.read',
				offline_access: ''
			},

			scope_delim: ' ',

			login: function(p) {
				if (p.qs.response_type === 'code') {
					// Let's set this to an offline access to return a refresh_token
					p.qs.access_type = 'offline_access';
				}
			},

			base: 'https://www.graph.microsoft.com/v1.0/',

			get: {
				me: 'me'
			},

			xhr: function(p) {
				if (p.method === 'post' || p.method === 'put') {
					toJSON(p);
				}
				else if (p.method === 'patch') {
					hello.utils.extend(p.query, p.data);
					p.data = null;
				}

				return true;
			},

			// Don't even try submitting via form.
			// This means no POST operations in <=IE9
			form: false
		}
	});
})(hello);

class OnedriveDB {
    constructor(appID, appName) {
        this.appID = appID;
        this.appName = appName;
        this.initalized = false;
    }

    async init() {
        return new Promise((resolve, reject) => {
            // Set up access to onedrive
            hello.init({
                aad: this.appID
            }, {
                redirect_uri: '../redirect.html',
                scope: 'User.Read, Files.ReadWrite.AppFolder'
            });

            this.getToken().then(token => {
                this.token = token;

                // Make sure the folder exists
                this.api_post(`me/drive/special/approot/children`, {name: this.appName, folder: {}}).then(data => {
                    this.initalized = true;

                    // Load data from all files
                    this.load().then(values => resolve(values), error => reject(error));
                });
            }, error => reject(error));
        });      
    }

    // Gets an access token
    // Logs the user in if he is not already logged in 
    async getToken() {
        return new Promise((resolve, reject) => {
            const response = hello.getAuthResponse('aad');
            if(response === null) {
                hello('aad').login({
                    display: 'popup'
                }).then(() => {
                    resolve(hello.getAuthResponse('aad').access_token);
                }, e => {
                    reject(e.error);
                });
            } else {
                resolve(response.access_token);
            }
        });
    }

    // Loads the content of all files in the app folder
    // Throws if atleast one load failed or the connection has not been inizalized
    // Returns the loaded data but also stores it in the class under .data
    async load() {
        return new Promise((resolve, reject) => {
            if(!this.initalized) 
                reject(new Error("Connection to OneDrive not inialized. Await .init() before doing anything else"));

            this.api_get(`me/drive/special/approot:/${this.appName}:/children`).then(data => {
                const promises = [];
                for(let item of data.value) {
                    promises.push(this.readFile(item.name));
                }

                Promise.all(promises).then(values => {
                    const result = {};

                    for(let value of values) {
                        result[value.name] = value.value;
                    }

                    this.data = result;

                    resolve(result);
                }, error => reject(error));
            }, error => reject(error));
        });
    }

    async save() {
        return new Promise((resolve, reject) => {
            this.deleteFiles().then(() => {
                const promises = [];
                let counter = 0;
    
                Object.keys(this.data).forEach(key => {
                    promises.push(this.createFile(counter + '.json', {
                        name: key,
                        value: this.data[key]
                    }));
                    counter++;
                });

                Promise.all(promises).then(values => {
                    resolve(values);
                }, error => reject(error));
            }, error => reject(error));
        });
    }

    async deleteFiles() {
        return new Promise((resolve, reject) => {
            const promises = [];

            this.api_get(`me/drive/special/approot:/${this.appName}:/children`).then(data => {
                const promises = [];
                for(let item of data.value) {
                    promises.push(this.deleteFile(item.name));
                }

                Promise.all(promises).then(values => {
                    resolve(values);
                }, error => reject(error));
            }, error => reject(error));
        });
    }

    async deleteFile(fileName) {
        return new Promise((resolve, reject) => {
            this.api_delete(`me/drive/special/approot:/${this.appName}/${fileName}:/`).then(result => resolve(result), error => reject(error));
        });
    }

    async readFile(fileName) {
        return new Promise((resolve, reject) => {   
            this.api_get(`me/drive/special/approot:/${this.appName}/${fileName}:/content`).then(data => {
                resolve(data);
            }, error => reject(error));
        });
    }

    async createFile(fileName, content) {
        return new Promise((resolve, reject) => {   
            this.api_put(`me/drive/special/approot:/${this.appName}/${fileName}:/content`, JSON.stringify(content)).then(data => {
                resolve(data);
            }, error => reject(error));
        });
    }

    async api_get(path) {
        return new Promise((resolve, reject) => {        
            const oReq = new XMLHttpRequest();

            oReq.onload = function () {
                if (this.status >= 200 && this.status < 300) {
                    resolve(JSON.parse(oReq.response));
                } else {
                    reject({
                        status: this.status,
                        statusText: oReq.statusText
                    });
                }
            };
            oReq.onerror = function () {
                reject({
                    status: this.status,
                    statusText: oReq.statusText
                });
            };

            oReq.open("GET", "https://graph.microsoft.com/v1.0/" + path, true);
            oReq.setRequestHeader('Authorization', 'Bearer ' + this.token);
            oReq.send();
        });
    }

    async api_post(path, body) {
        return new Promise((resolve, reject) => {        
            const oReq = new XMLHttpRequest();

            oReq.onload = function () {
                if (this.status >= 200 && this.status < 300) {
                    resolve(oReq.response);
                } else {
                    reject({
                        status: this.status,
                        statusText: oReq.statusText
                    });
                }
            };
            oReq.onerror = function () {
                reject({
                    status: this.status,
                    statusText: oReq.statusText
                });
            };

            oReq.open("POST", "https://graph.microsoft.com/v1.0/" + path, true);
            oReq.setRequestHeader('Authorization', 'Bearer ' + this.token);
            oReq.setRequestHeader("Content-Type", "application/json");
            oReq.send(JSON.stringify(body));
        });
    }

    async api_put(path, body) {
        return new Promise((resolve, reject) => {        
            const oReq = new XMLHttpRequest();

            oReq.onload = function () {
                if (this.status >= 200 && this.status < 300) {
                    resolve(oReq.response);
                } else {
                    reject({
                        status: this.status,
                        statusText: oReq.statusText
                    });
                }
            };
            oReq.onerror = function () {
                reject({
                    status: this.status,
                    statusText: oReq.statusText
                });
            };

            oReq.open("PUT", "https://graph.microsoft.com/v1.0/" + path, true);
            oReq.setRequestHeader('Authorization', 'Bearer ' + this.token);
            oReq.setRequestHeader("Content-Type", "text/plain");
            oReq.send(body);
        });
    }

    async api_delete(path) {
        return new Promise((resolve, reject) => {        
            const oReq = new XMLHttpRequest();

            oReq.onload = function () {
                if (this.status >= 200 && this.status < 300) {
                    resolve(oReq.response);
                } else {
                    reject({
                        status: this.status,
                        statusText: oReq.statusText
                    });
                }
            };
            oReq.onerror = function () {
                reject({
                    status: this.status,
                    statusText: oReq.statusText
                });
            };

            oReq.open("DELETE", "https://graph.microsoft.com/v1.0/" + path, true);
            oReq.setRequestHeader('Authorization', 'Bearer ' + this.token);
            oReq.send();
        });
    }
}

