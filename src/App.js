/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

// This sample uses an open source OAuth 2.0 library that is compatible with the Azure AD v2.0 endpoint.
// Microsoft does not provide fixes or direct support for this library.
// Refer to the library’s repository to file issues or for other support.
// For more information about auth libraries see: https://docs.microsoft.com/azure/active-directory/active-directory-v2-libraries
// Library repo: https://github.com/MrSwitch/hello.js

import React, { Component } from 'react';
import hello from 'hellojs';
window.hello = hello;
import GraphSdkHelper from './helpers/GraphSdkHelper';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import DetailsListExample from './component-examples/DetailsList';
import { applicationId, redirectUri } from './helpers/config';

export default class App extends Component {
  constructor(props) {
    super(props);
    
    // Initialize the auth network.
    hello.init({
      aad: {
        name: 'Azure Active Directory',	
        oauth: {
          version: 2,
          auth: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
        },
        form: false
      }
    });
    
    // Initialize the Graph SDK helper and save it in the window object.
    this.sdkHelper = new GraphSdkHelper({ login: this.login.bind(this) });
    window.sdkHelper = this.sdkHelper;

    // Set the isAuthenticated prop and the (empty) Fabric example selection. 
    this.state = {
      isAuthenticated: !!hello('aad').getAuthResponse(),
      example: ''
    };
  }

  // Get the user's display name.
  componentWillMount() {
    if (this.state.isAuthenticated) {
      this.sdkHelper.getMe((err, me) => {
        if (!err) {
          this.setState({
            displayName: `Hello ${me.displayName}!`
          });
        }
      });
    }
  }

  // Sign the user into Azure AD. HelloJS stores token info in localStorage.hello.
  login() {

    // Initialize the auth request.
    hello.init( {
      aad: applicationId
      }, {
      redirect_uri: redirectUri,
      scope: 'user.readbasic.all+mail.send+files.read'
    });

    hello.login('aad', { 
      display: 'page',
      state: 'abcd'
    });
  }

  // Sign the user out of the session.
  logout() { 
    hello('aad').logout();
    this.setState({ 
      isAuthenticated: false,
      example: '',
      displayName: ''
    });
  }

  render() {
    return (
      <div>
        <div>
        {
          
          // Show the command bar with the Sign in or Sign out button.
          <CommandBar
            items={[
              {
                key: 'details-list-example',
                name: 'OneDrive Files List',
                disabled: !this.state.isAuthenticated,
                ariaLabel: 'Choose a component example to render in the page',
                onClick: () => { this.setState({ example: 'details-list-example' }) }
              }
            ]}
            farItems={[
              {
                key: 'display-name',
                name: this.state.displayName
              },
              {
                key: 'log-in-out=button',
                name: this.state.isAuthenticated ? 'Sign out' : 'Sign in',
                onClick: this.state.isAuthenticated ? this.logout.bind(this) : this.login.bind(this)
              }
            ]} />
        }
        </div>
        <div className="ms-font-m">
          <div>
            <h2>OneDrive Files</h2>
            {

              (!this.state.isAuthenticated || this.state.example === '') &&
              <div>
              <p>To get started, sign in.</p>
              </div>
            }
          </div>
          <br />
          {
            
            // Show the selected fabric component example.
            this.state.isAuthenticated &&
              <div>
              {
                this.state.example === 'details-list-example' &&
                <DetailsListExample />
              }
              </div>
          }
          <br />
        </div>
      </div>
    );
  }
}
