import * as React from 'react';
import styles from './RegistrationForm.module.scss';
import { IRegistrationFormProps } from './IRegistrationFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Dialog, DialogType, Dropdown, MessageBar, MessageBarType, Modal, Spinner, SpinnerSize } from 'office-ui-fabric-react';
import * as pnp from "sp-pnp-js";


export default class RegistrationForm extends React.Component<IRegistrationFormProps, {}> {
  state = {
    isUserPresent: null,
    isError: null,
    showLoader: false,
    ID: null,
    currnetUserDetails:null,
    messageBox: {
      show: false,
      type: null,
      message: ''
    },
    modalPopup: {
      text: "",
      show: false
    },
    approvalOptions: ["Approved", "Rejected", "In Progress"],
    form: {
      title: {
        label: "Title*",
        elementConfig: {
          type: "input",
          dataType: "text",
          placeholder: "enter title"
        },
        value: '',
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      firstname: {
        label: "First Name*",
        elementConfig: {
          type: "input",
          dataType: "text",
          placeholder: "Enter First Name"
        },
        value: '',
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      middlename: {
        label: "Middle Name",
        elementConfig: {
          type: "input",
          dataType: "text",
          placeholder: "Enter Middle Name"
        },
        value: '',
        rules: {
          required: false
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      lastname: {
        label: "Last Name*",
        elementConfig: {
          type: "input",
          dataType: "text",
          placeholder: "Enter Last Name"
        },
        value: '',
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      phoneno: {
        label: "Phone No*",
        elementConfig: {
          type: "input",
          dataType: "number",
          placeholder: "Enter Contact Number"
        },
        value: '',
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      companyname: {
        label: "Company Name*",
        elementConfig: {
          type: "input",
          dataType: "text",
          placeholder: "Enter Company Name"
        },
        value: '',
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      companyaddress: {
        label: "Company Address*",
        elementConfig: {
          type: "input",
          dataType: "text",
          placeholder: "Enter Company Address"
        },
        value: '',
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      divisionDepartment: {
        label: "Division/Department*",
        elementConfig: {
          type: "input",
          dataType: "text",
          placeholder: "Enter Division/Department"
        },
        value: '',
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      country: {
        label: "Country*",
        elementConfig: {
          type: "input",
          dataType: "text",
          placeholder: "Enter Country"
        },
        value: '',
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      StateProvince: {
        label: "State Provience*",
        elementConfig: {
          type: "input",
          dataType: "text",
          placeholder: "Enter State Provience"
        },
        value: '',
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      city: {
        label: "City*",
        elementConfig: {
          type: "input",
          dataType: "text",
          placeholder: "Enter City"
        },
        value: '',
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      zipPostalCode: {
        label: "Zip Code*",
        elementConfig: {
          type: "input",
          dataType: "number",
          placeholder: "Enter Zip Code"
        },
        value: '',
        rules: {
          required: true,
          length : 6 
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      day1attendence: {
        label: "Day 1 Attendance*",
        elementConfig: {
          type: "dropdown",
          dataType: "select",
          placeholder: "Select Attendence"
        },
        value: "Yes",
        options: [{ key: 'Yes', text: 'Yes' },
        { key: 'No', text: 'No' }],
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      day2attendence: {
        label: "Day 2 Attendance*",
        elementConfig: {
          type: "dropdown",
          dataType: "select",
          placeholder: "Select Attendence"
        },
        value: "Yes",
        options: [{ key: 'Yes', text: 'Yes' },
        { key: 'No', text: 'No' },],
        rules: {
          required: true
        },
        error: {
          isValid: true,
          message: ""
        }
      },
      submitBtn: {
        label: "Submit to Hitachi",
        elementConfig: {
          type: "button",
          dataType: "button",
          placeholder: "Submit to Hitachi"
        },
        value: true, //disable or enable
        rules: {
          required: false
        },
        error: {
          isValid: true,
          message: ""
        }
      }
    }
  }
  public validateForm = (updatedForm, id) => {
    let isValid = true;
    updatedForm[id].error.isValid = true;
    updatedForm[id].error.message = "";
    if (updatedForm[id].rules["required"] && updatedForm[id].value === "") {
      updatedForm[id].error.message = "Required Field";
      updatedForm[id].error.isValid = false && isValid;
    }
    if (updatedForm[id].rules["length"] && updatedForm[id].value.length !== updatedForm[id].rules["length"]) {
      updatedForm[id].error.message = "length should be "+updatedForm[id].rules["length"];
      updatedForm[id].error.isValid = false  && isValid;
    }
  }
  public spLoggedInUserDetails(ctx: any): Promise<any> {
    try {
      const web = new pnp.Web(ctx.pageContext.site.absoluteUrl);
      return web.currentUser.get();
    } catch (error) {
      console.log("Error in spLoggedInUserDetails : " + error);
    }
  }

  private checkUserIfPresent(): Promise<any> {
    try {
      //console.log('in check user present');
      let userDetails = this.spLoggedInUserDetails(this.props.context);
      return userDetails.then((params: any) => {
        //console.log(params);
        this.setState({currnetUserDetails:params});
        return pnp.sp.web.lists.getById(this.props.lists).items
          .filter(`EventUserId/EMail eq '${encodeURIComponent(params.Email)}'`)
          .get().then((result) => {
            this.setState({ showLoader: false });
            if (result.length > 0) {
              //console.log('user present');
              this.setState({ isUserPresent: true });
              this.setFormData(result[0]);
            } else {
              //console.log('user not present');
              this.setState({ isUserPresent: false });
            }
          }).catch((er) => {
            console.log(er);
            this.setState({ showLoader: false });
            this.setState({ isError: er });
          });
      });
    } catch (error) {
      this.setState({ showLoader: false, isError: error });
      console.log("Error in loadUserDetails : ", error);
    }
  }
  public setFormData(userData: any) {
    //console.log(userData);
    let updateForm = { ...this.state.form };
    updateForm["title"].value = userData["Title"];
    updateForm["firstname"].value = userData["FirstName"];
    updateForm["middlename"].value = userData["MiddleName"];
    updateForm["lastname"].value = userData["LastName"];
    updateForm["phoneno"].value = userData["PhoneNo"];
    updateForm["companyname"].value = userData["CompanyName"];
    updateForm["companyaddress"].value = userData["CompanyAddress"];
    updateForm["divisionDepartment"].value = userData["Division_Department"];
    updateForm["city"].value = userData["City"];
    updateForm["country"].value = userData["Country"];
    updateForm["StateProvince"].value = userData["StateProvince"];
    updateForm["zipPostalCode"].value = userData["ZipPostalCode"];
    updateForm["day1attendence"].value = userData["Day1Attendance"] == true ? "Yes" : "No";
    updateForm["day2attendence"].value = userData["Day2Attendance"] == true ? "Yes" : "No";
    let modal = null;
    if (userData["ApplicationStatus"] === this.state.approvalOptions[0]) {
      modal = { text: "Please wait redirecting...", show: true };
      updateForm["submitBtn"].value = false;
      if (this.props.homePageUrl) {
        location.href = this.props.homePageUrl;
      }
    } else if (userData["ApplicationStatus"] === this.state.approvalOptions[1]) {
      modal = { text: "Request Rejected", show: true };
      updateForm["submitBtn"].value = false;
    } else if (userData["ApplicationStatus"] === this.state.approvalOptions[2]) {
      modal = { text: "Request in progress", show: true };
      updateForm["submitBtn"].value = false;
    }else{
      modal = { text: "Request in progress", show: false };
      updateForm["submitBtn"].value = true;
    }
    this.setState({ form: updateForm, ID: userData["ID"], modalPopup: modal });
  }
  private closeModal = () => {
    let updatedModal = { ...this.state.modalPopup };
    updatedModal.show = false;
    this.setState({ modalPopup: updatedModal });
  }
  private createItem(): void {
    this.setState({ showLoader: true });
    pnp.sp.web.lists.getById(this.props.lists).items.add({
      Title: this.state.form.title.value,
      FirstName: this.state.form.firstname.value,
      MiddleName: this.state.form.middlename.value,
      LastName: this.state.form.lastname.value,
      PhoneNo: this.state.form.phoneno.value,
      CompanyName: this.state.form.companyname.value,
      CompanyAddress: this.state.form.companyaddress.value,
      Division_Department: this.state.form.divisionDepartment.value,
      City: this.state.form.city.value,
      Country: this.state.form.country.value,
      StateProvince: this.state.form.StateProvince.value,
      ZipPostalCode: this.state.form.zipPostalCode.value,
      Day1Attendance: this.state.form.day1attendence.value === "Yes" ? true : false,
      Day2Attendance: this.state.form.day2attendence.value === "Yes" ? true : false,
      ApplicationStatus: this.state.approvalOptions[2],
      EventUserId: this.state.currnetUserDetails.Id,
    }).then((res) => {
      this.setState({ showLoader: false });
      let messageBox = { ...this.state.messageBox };
      messageBox.message = "Sucess";
      messageBox.type = MessageBarType.success;
      messageBox.show = true;
      this.setState({ messageBox: messageBox });
    }).catch(err => {
      this.setState({ showLoader: false });
      let messageBox = { ...this.state.messageBox };
      messageBox.message = "Internal Server error";
      messageBox.type = MessageBarType.error;
      messageBox.show = true;
      this.setState({ messageBox: messageBox });
      console.log(err);
    });
  }
  private updateItem(): void {
    this.setState({ showLoader: true });
    pnp.sp.web.lists.getById(this.props.lists).items.getById(this.state.ID).update({
      Title: this.state.form.title.value,
      FirstName: this.state.form.firstname.value,
      MiddleName: this.state.form.middlename.value,
      LastName: this.state.form.lastname.value,
      PhoneNo: this.state.form.phoneno.value,
      CompanyName: this.state.form.companyname.value,
      CompanyAddress: this.state.form.companyaddress.value,
      Division_Department: this.state.form.divisionDepartment.value,
      City: this.state.form.city.value,
      Country: this.state.form.country.value,
      StateProvince: this.state.form.StateProvince.value,
      ZipPostalCode: this.state.form.zipPostalCode.value,
      Day1Attendance: this.state.form.day1attendence.value === "Yes" ? true : false,
      Day2Attendance: this.state.form.day2attendence.value === "Yes" ? true : false,
      ApplicationStatus: this.state.approvalOptions[2],
      EventUserId: this.state.currnetUserDetails.Id,
    }).then((res) => {
      this.setState({ showLoader: false });
      let messageBox = { ...this.state.messageBox };
      messageBox.message = "Sucess";
      messageBox.type = MessageBarType.success;
      messageBox.show = true;
      this.setState({ messageBox: messageBox });
    }).catch(err => {
      this.setState({ showLoader: false });
      let messageBox = { ...this.state.messageBox };
      messageBox.message = "Internal Server error";
      messageBox.type = MessageBarType.error;
      messageBox.show = true;
      this.setState({ messageBox: messageBox });
      console.log(err);
    });
  }

  public componentDidMount() {
    pnp.setup({
      spfxContext: this.props.context
    });
    this.setState({ showLoader: true });
    this.checkUserIfPresent();
  }
  public onChangeHandler = (ev, id) => {
    let updatedForm = {};
    updatedForm = { ...this.state.form };
    updatedForm[id] = { ...this.state.form[id] };
    updatedForm[id].value = ev.target.value;
    updatedForm[id].error = { ...this.state.form[id].error }
    this.validateForm(updatedForm, id);
    this.setState({ form: updatedForm });
  }
  public onChangeDropdownHandler = (ev, option, id) => {
    let updatedForm = {};
    updatedForm = { ...this.state.form };
    updatedForm[id] = { ...this.state.form[id] };
    updatedForm[id].value = option.key;
    updatedForm[id].error = { ...this.state.form[id].error }
    this.validateForm(updatedForm, id);
    this.setState({ form: updatedForm });
  }

  public submitFormHandler = (event): void => {
    event.preventDefault();
    let isValid = true;
    let updatedForm = { ...this.state.form };
    for (const input in updatedForm) {
      if (Object.prototype.hasOwnProperty.call(updatedForm, input)) {
        this.validateForm(updatedForm, input);
        isValid = isValid && updatedForm[input].error.isValid;
      }
    }
    this.setState({ form: updatedForm });

    if (isValid) {
      //console.log('submit form');
      if (this.state.ID) {
        this.updateItem();
      } else {
        this.createItem();
      }
    } else {
      console.log("form not valid");
      let messageBox = { ...this.state.messageBox };
      messageBox.message = "Invalid Form";
      messageBox.type = MessageBarType.error;
      messageBox.show = true;
      this.setState({ messageBox: messageBox });
    }
  }

  public render(): React.ReactElement<IRegistrationFormProps> {
    //console.log(this.props.startDate, this.props.endDate);
    let formGroups = [
      { id: "companyname" },
      { id: "companyaddress" },
      { id: "country" },
      { id: "zipPostalCode" }
    ];

    let companyInformation1 = formGroups.map(item => (
      <div className={styles["form-group"]}>
        <label htmlFor={item.id}>{this.state.form[item.id].label}</label>
        <input value={this.state.form[item.id].value} onChange={(evnt) => this.onChangeHandler(evnt, item.id)} type={this.state.form[item.id].elementConfig.dataType} className={styles["form-control"]} placeholder={this.state.form[item.id].elementConfig.placeholder}></input>
        <small className={[styles["form-text"], styles["text-muted"], styles["custom-error"]].join(" ")} >{this.state.form[item.id].error.message}</small>
      </div>
    ));

    formGroups = [
      { id: "divisionDepartment" },
      { id: "city" },
      { id: "StateProvince" }
    ];

    let companyInformation2 = formGroups.map(item => (
      <div className={styles["form-group"]}>
        <label htmlFor={item.id}>{this.state.form[item.id].label}</label>
        <input value={this.state.form[item.id].value} onChange={(evnt) => this.onChangeHandler(evnt, item.id)} type={this.state.form[item.id].elementConfig.dataType} className={styles["form-control"]} placeholder={this.state.form[item.id].elementConfig.placeholder}></input>
        <small className={[styles["form-text"], styles["text-muted"], styles["custom-error"]].join(" ")} >{this.state.form[item.id].error.message}</small>
      </div>
    ));

    formGroups = [
      { id: "firstname" },
      { id: "lastname" },
      { id: "title" },
    ];

    let personalInformation1 = formGroups.map(item => (
      <div className={styles["form-group"]}>
        <label htmlFor={item.id}>{this.state.form[item.id].label}</label>
        <input value={this.state.form[item.id].value} onChange={(evnt) => this.onChangeHandler(evnt, item.id)} type={this.state.form[item.id].elementConfig.dataType} className={styles["form-control"]} placeholder={this.state.form[item.id].elementConfig.placeholder}></input>
        <small className={[styles["form-text"], styles["text-muted"], styles["custom-error"]].join(" ")} >{this.state.form[item.id].error.message}</small>
      </div>
    ));

    formGroups = [
      { id: "middlename" },
      { id: "phoneno" }
    ];

    let personalInformation2 = formGroups.map(item => (
      <div className={styles["form-group"]}>
        <label htmlFor={item.id}>{this.state.form[item.id].label}</label>
        <input value={this.state.form[item.id].value} onChange={(evnt) => this.onChangeHandler(evnt, item.id)} type={this.state.form[item.id].elementConfig.dataType} className={styles["form-control"]} placeholder={this.state.form[item.id].elementConfig.placeholder}></input>
        <small className={[styles["form-text"], styles["text-muted"], styles["custom-error"]].join(" ")} >{this.state.form[item.id].error.message}</small>
      </div>
    ));

    let form =
      <div>
        {
          this.state.modalPopup.show ?
            <Dialog
              isOpen={this.state.modalPopup.show}
              type={DialogType.normal}
              // onDismiss={this.closeModal.bind(this)}
              // title={this.state.modalPopup.text}
              isBlocking={true}
              closeButtonAriaLabel='Close'
            >
              <h1>{this.state.modalPopup.text}</h1>
            </Dialog> : ""
        }
        {this.state.messageBox.show ?
          <MessageBar
            messageBarType={this.state.messageBox.type}
            isMultiline={false}
            onDismiss={() => this.setState({
              messageBox: {
                show: false,
                type: null,
                message: ''
              }
            })}
            dismissButtonAriaLabel="Close"
          >
            {this.state.messageBox.message}
          </MessageBar> : ""
        }
        <div className={[styles.row, styles.customROW].join(" ")}>
          <div className={styles.col}>
            <div className={styles["form-group"]}>
              <span className={styles.lbcustom} >Event Information</span>
            </div>
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles["col-md-6"]}>
            <div className={styles["form-group"]}>
              <label htmlFor="StartDate">Start date : {this.props.startDate ? (new Date(this.props.startDate.displayValue).toDateString()) : ""}</label><br />
              <label htmlFor="ContactEmail">{this.props.additionalInformation ? this.props.additionalInformation : ""}</label><br />
            </div>
          </div>
          <div className={styles["col-md-6"]}>
            <div className={styles["form-group"]}>
              <label htmlFor="inputEmail4">End date : {this.props.endDate ? (new Date(this.props.endDate.displayValue)).toDateString() : ""}</label>
            </div>
          </div>
        </div>
        <div className={[styles.row, styles.customROW].join(" ")}>
          <div className={styles.col}>
            <div className={styles["form-group"]}>
              <span className={styles.lbcustom} >Personal Information</span>
            </div>
          </div>
        </div>


        <form onSubmit={this.submitFormHandler}>
          <div className={styles.row}>
            <div className={styles["col-md-6"]}>
              {personalInformation1}
            </div>
            <div className={styles["col-md-6"]}>
              {personalInformation2}
            </div>
          </div>
          <div className={[styles.row, styles.customROW].join(" ")}>
            <div className={styles.col}>
              <div className={styles["form-group"]}>
                <span className={styles.lbcustom} >Company Information</span>
              </div>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles["col-md-6"]}>
              {companyInformation1}
            </div>
            <div className={styles["col-md-6"]}>
              {companyInformation2}
            </div>
          </div>
          <div className={[styles.row, styles.customROW].join(" ")}>
            <div className={styles.col}>
              <div className={styles["form-group"]}>
                <span className={styles.lbcustom} >Conference Information</span>
              </div>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles["col-md-6"]}>
              <div className={styles["form-group"]}>
                <Dropdown
                  label={this.state.form.day1attendence.label}
                  selectedKey={this.state.form.day1attendence.value}
                  options={this.state.form.day1attendence.options}
                  onChange={(ev, choice) => this.onChangeDropdownHandler(ev, choice, "day1attendence")}
                />
                <small className={[styles["form-text"], styles["text-muted"], styles["custom-error"]].join(" ")} >{this.state.form.day1attendence.error.message}</small>
              </div>
            </div>
            <div className={styles["col-md-6"]}>
              <div className={styles["form-group"]}>
                <Dropdown
                  label={this.state.form.day2attendence.label}
                  selectedKey={this.state.form.day2attendence.value}
                  options={this.state.form.day2attendence.options}
                  onChange={(ev, choice) => this.onChangeDropdownHandler(ev, choice, "day2attendence")}
                />
                <small className={[styles["form-text"], styles["text-muted"], styles["custom-error"]].join(" ")} >{this.state.form.day2attendence.error.message}</small>
              </div>
            </div>
          </div>
          <div className={[styles.contbtn].join(" ")}>
            <div className={[styles["col-md-12"], styles["text-center"]].join(" ")}>
              <button disabled={!this.state.form.submitBtn.value} type="submit" className={[styles.btn, styles.btnsubmit, styles["btn-primary"]].join(" ")}>{this.state.form.submitBtn.label}</button>
            </div>
          </div>
        </form>
      </div>

    if (this.state.isError === null) {
      this.props.context.statusRenderer.clearError(this.props.domElement);
    } else {
      this.props.context.statusRenderer.renderError(this.props.domElement, this.state.isError);
    }

    return (
      <div className={styles.registrationForm}>
        {
          this.state.showLoader ?
            <Spinner size={SpinnerSize.large} label="Form Loading" /> :
            form
        }
      </div>
    );
  }
}
