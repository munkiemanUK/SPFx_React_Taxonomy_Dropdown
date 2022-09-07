import * as React from 'react';
import styles from './ReactTaxonomyDropdown.module.scss';
import { IReactTaxonomyDropdownProps } from './IReactTaxonomyDropdownProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { spfi } from '@pnp/sp';    
import { getGUID } from "@pnp/common"; 

export default class ReactTaxonomyDropdown extends React.Component<IReactTaxonomyDropdownProps, {}> {

  public componentWillMount() {
    taxonomy.getDefaultSiteCollectionTermStore()
      .getTermSetById('b8ab6fd6-38e9-4ac8-ba86-d07bf4ec530f')
      .terms.get().then(
        Allterms => {
          console.log(Allterms);
          this.setState({terms: Allterms})
        }
      )
  }

  public render(): React.ReactElement<IReactTaxonomyDropdownProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.reactTaxonomyDropdown} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <select id="dropdown" onChange={this.handleDropdownChange}>
          { this.state.terms.map(term => {
              return <option value={term.Name.toLowerCase()}>{term.Name}</option>
            })
          }
        </select>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }
}
