import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import {SPHttpClient,SPHttpClientResponse,ISPHttpClientOptions} from "@microsoft/sp-http";

import {List,Lists} from "./helper"



import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

import 'bootstrap/dist/css/bootstrap.min.css';

//import 'bootstrap/dist/js/bootstrap.bundle.min.js';

import 'bootstrap/dist/js/bootstrap.bundle.min.js';
//import { response } from 'express';


export interface IHelloWorldWebPartProps {
  description: string;
}


export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
        The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
          <ul class="${styles.links}">
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>
      </div>

      <!-- this is comment --->



      <div class="container mt-5">
  <h2>Personal Information Form</h2>


 <!-- <form onsubmit = "submitFormData()"> -->
  
    <div class="form-group mb-3">
      <label for="idNumber">ID</label>
      <input type="text" class="form-control" id="idNumber" placeholder="Enter your ID number">
      <span><button id = "fetch">Fetch Record</button></span>
    </div>
  
    <div class="form-group mb-3">
      <label for="firstName">First Name:</label>
      <input type="text" class="form-control" id="firstName" placeholder="Enter your first name">
    </div>
	
	<div class="form-group mb-3">
      <label for="lastName">Last Name:</label>
      <input type="text" class="form-control" id="lastName" placeholder="Enter your last name">
    </div>

   
    <div class="form-group mb-3">
      <label for="email">Email:</label>
      <input type="email" class="form-control mb-10" id="email" placeholder="Enter your email">
    </div>
	
	<div class="form-group mb-3">
      <label for="lastName">Location:</label>
      <input type="text" class="form-control" id="location" placeholder="Enter your location">
    </div>
	
	<select id="select" class="form-select mb-3" aria-label="Default select example">
  <option selected>status</option>
  <option value="1">Approved</option>
  <option value="2">pending</option>
  <option value="3">Rejected</option>
</select>
	

    <button id="form1" type="submit" class="btn btn-primary ">Submit</button>
<!--  </form>  -->


</div>

<p></p>
<p></p>

<br/>
<br/>



<button type="button"  class="btn btn-primary" id="btn1">Read</button>
<button type="button" class="btn btn-secondary" id = "btn2">Update</button>
<button type="button" class="btn btn-success">Create</button>
<button type="button" class="btn btn-danger">Delete</button>
<button type="button" class="btn btn-warning" id="btn5">Clear below Data</button>

<br/>
<br/>
      <div id = "data">
       
      </div>



    </section>`;

     //this.getListData();
     const readButton = this.domElement.querySelector("#btn1") as HTMLButtonElement;
     readButton.addEventListener("click",()=>{

      this.temp();
      

     });

     const clearButton = this.domElement.querySelector("#btn5") as HTMLButtonElement;
     clearButton.addEventListener("click",()=>{
        this.temp1();
     });

     // now submitting form data to the sharepoint list:-

     const form1 = this .domElement.querySelector("#form1") as HTMLButtonElement;

     form1.addEventListener('click',()=>{
      this.submitFormData();
     })

     // getlist data based on id:-
     const fetchId = this.domElement.querySelector("#fetch") as HTMLButtonElement;
     fetchId.addEventListener("click",()=>{
      this.getListItemById();
     })



     // update sharepoint list data:-

     const update = this.domElement.querySelector("#btn2") as HTMLButtonElement;

     update.addEventListener("click",()=>{
      this.updateListItem();
     })


  

     
    
  }


  // this is the complete code for getting data from sharepoint list using rest api:-

  // erasing the view item:-

  private temp1():void{
     const data1 = this.domElement.querySelector("#data") as HTMLDivElement;
     data1.innerHTML = "";
  }

  private temp():void{

    this.getResponse();
  
  }

  private renderList(value:List[]):void{

    console.log("now we are at the renderList function:");

    let html:string = `<table class= " ${styles.table1}">`

     html+= `

      <tr>

      <th class = "${styles.th1}">Id</th>
      <th class = "${styles.th1}">FirstName</th>
      <th class = "${styles.th1}">SecondName</th>
      <th class = "${styles.th1}">Email</th>
      <th class = "${styles.th1}">Location</th>
      <th class = "${styles.th1}">Status</th>
      
      </tr>

      `

      
      

    value.forEach((item:List)=>{

      
      
      html+= `
        

        <tr>
        <td class = " ${styles.td1}">${item.Title}</td>
        
        <td  class = " ${styles.td1}">${item.FirstName}</td>
        <td  class = " ${styles.td1}">${item.SecondName}</td>
        <td  class = " ${styles.td1}">${item.Email}</td>
        <td  class = " ${styles.td1}">${item.Location}</td>
        <td  class = " ${styles.td1}">${item.Status}</td>
        
        </tr>
        
       
      `
    })

    html+= `</table>`

    const data = this.domElement.querySelector("#data") as HTMLDivElement;
    console.log(html);

    data.innerHTML = html;

  }
  
  private getResponse():void{
     
    this.getListData()
    .then((response)=>{
      console.log("the getresponse is working:");
       this.renderList(response.value);
    })
    .catch((error)=>{
      console.log("this is error: ",error);
    })

  }
  
 private getListData():Promise<Lists>{
    
  return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Customer')/items",SPHttpClient.configurations.v1)
  .then((response:SPHttpClientResponse)=>{

    if(response.status === 200){
      console.log("list data is fatched:");
     // console.log("response.json() value is =  ", response.json());
      

       return response.json();
     
    }
    else{
      console.log("not fatched",response.status);

      return response.status;
    }
    
  })
  .catch((error)=>{
    console.log(error);
  })
 }


// using a post request we are creating a new record in a list:-

private submitFormData():void {

//   .options[statusSelect.selectedIndex];
// const selectedText = selectedOption.innerText;

  const Title = (this.domElement.querySelector("#idNumber") as HTMLInputElement).value;
  const firstName = (this.domElement.querySelector("#firstName") as HTMLInputElement).value;
  const lasttName = (this.domElement.querySelector("#lastName") as HTMLInputElement).value;
  const email = (this.domElement.querySelector("#email") as HTMLInputElement).value;
  const location = (this.domElement.querySelector("#location") as HTMLInputElement).value;
  // const status = (this.domElement.querySelector("#select") as HTMLSelectElement);
  // console.log(status);

  const statusSelect = this.domElement.querySelector("#select") as HTMLSelectElement;
const selectedOption = statusSelect.options[statusSelect.selectedIndex];
const status = selectedOption.innerText;

console.log('Selected Text:', status);

  

  console.log(Title,firstName,lasttName,email,location,status);

  const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Customer')/items`;
 
  const itemData = {
   // '__metadata': { 'type': 'SP.Data.CustomerListItem' },
    'Title': Title,
    'FirstName':firstName,
    'SecondName':lasttName,
    'Email':email,
    'Location':location,
    'Status':status

   
  };
  // {
  //   "headers": {
  //     'Accept': 'application/json;odata=nometadata',
  //     'Content-Type': 'application/json;odata=nometadata',
  //     'odata-version': ''
  //   },
  //   "body": JSON.stringify(itemData)
  // }

  const config:ISPHttpClientOptions = {
    "body":JSON.stringify(itemData)
  }

  this.context.spHttpClient.post(listUrl, SPHttpClient.configurations.v1, config)
  .then((response: SPHttpClientResponse) => {
    if (response.ok) {
      console.log('Item created successfully');
      alert("list item created successfully:");
      // Add any additional logic after successful creation
    } else {
      console.error(`Error creating item: ${response.statusText}`);
      console.log("the list item are not created:");
    }
  })
  .catch((error) => {
    console.error('Error creating item', error);
  });
}





// fetching record based on specifi id :

private getListItemById(): void {

  const listName = 'Customer'; // Replace with your actual list name
  const itemId =  (this.domElement.querySelector("#idNumber") as HTMLInputElement).value;
  const endpointUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`;

  this.context.spHttpClient.get(endpointUrl, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => response.json())
    .then((data) => {
      console.log('Item data:', data);
      // Process the retrieved item data
      console.log("this is data.Title",data.Title);
      console.log("this is firstname:",data.FirstName);
      console.log("this is data.value",data.value);
      console.log(data.Status);

       (this.domElement.querySelector("#idNumber") as HTMLInputElement).value =data.Title;
       (this.domElement.querySelector("#firstName") as HTMLInputElement).value = data.FirstName;
       (this.domElement.querySelector("#lastName") as HTMLInputElement).value = data.SecondName;
       (this.domElement.querySelector("#email") as HTMLInputElement).value = data.Email;
       (this.domElement.querySelector("#location") as HTMLInputElement).value = data.Location;
      const new1 =  (this.domElement.querySelector("#select") as HTMLSelectElement);
      console.log(new1);
      console.log("before value update is :",new1.value);
     // console.log("the textcontent porperty",new1.textContent);
      //(this.domElement.querySelector("#select") as HTMLSelectElement) = data.Status
      // new1.options[new1.selectedIndex] = data.Status.;
      // console.log(new1.innerHTML);
      // console.log(new1.options);
      // console.log(new1.options[new1.selectedIndex].innerText);

      // const newstr:any = "0";

      // const index = data.Status - newstr;
       
      //  new1.value = new1.options[data.Status - index].innerText;
      //  console.log("the updated new1.value is ",new1.value);

       new1.value = (new1.options[new1.selectedIndex]).text;
       console.log("the updated new1.value is ",new1.value);
       

      // console.log(new1.selectedOptions);

      //  const statusSelect = this.domElement.querySelector("#select") as HTMLSelectElement;
      //    const selectedOption = statusSelect.options[statusSelect.selectedIndex];
      //  const status = selectedOption.innerText;
     
    })
    .catch((error) => {
      console.error('Error retrieving item', error);
    });
}



// update the sharepoint list Item.

private updateListItem():void{

  const Title = (this.domElement.querySelector("#idNumber") as HTMLInputElement).value;
  const firstName = (this.domElement.querySelector("#firstName") as HTMLInputElement).value;
  const lasttName = (this.domElement.querySelector("#lastName") as HTMLInputElement).value;
  const email = (this.domElement.querySelector("#email") as HTMLInputElement).value;
  const location = (this.domElement.querySelector("#location") as HTMLInputElement).value;
  

  const statusSelect = this.domElement.querySelector("#select") as HTMLSelectElement;
  const selectedOption = statusSelect.options[statusSelect.selectedIndex];
  const status = selectedOption.innerText;





  const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Customer')/items(${Title})`;
 
  const itemData = {
   // '__metadata': { 'type': 'SP.Data.CustomerListItem' },
    'Title': Title,
    'FirstName':firstName,
    'SecondName':lasttName,
    'Email':email,
    'Location':location,
    'Status':status

   
  };

  const header = {
    "X-HTTP-Method":"MERGE",
    "IF-MATCH": "*"
  }
  

  const config:ISPHttpClientOptions = {

    "body":JSON.stringify(itemData),
    "headers":header

  }

  this.context.spHttpClient.post(listUrl, SPHttpClient.configurations.v1, config)
  .then((response: SPHttpClientResponse) => {
    if (response.ok) {
      console.log('Item created successfully');
      alert("list item updated successfully:");
      // Add any additional logic after successful creation
    } else {
      console.error(`Error creating item: ${response.statusText}`);
      console.log("the list item are not created:");
      alert("list item is not updated.");
    }
  })
  .catch((error) => {
    console.error('Error creating item', error);
  });



}






  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
