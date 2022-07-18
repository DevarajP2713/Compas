import * as React from "react";
import Table from "@material-ui/core/Table";
import TableBody from "@material-ui/core/TableBody";
import TableCell from "@material-ui/core/TableCell";
import TableHead from "@material-ui/core/TableHead";
import TablePagination from "@material-ui/core/TablePagination";
import TableRow from "@material-ui/core/TableRow";
import TableSortLabel from "@material-ui/core/TableSortLabel";
import Toolbar from "@material-ui/core/Toolbar";
import Typography from "@material-ui/core/Typography";
import Paper from "@material-ui/core/Paper";
import Pagination from "@material-ui/lab/Pagination";
import { lighten } from "@material-ui/core/styles/colorManipulator";
import { TableContainer } from "@material-ui/core";
import classes from "./App.module.scss";
import Header from "./Header";
import TopFilter from "./TopFilter";
import Panel from "./Panel";
import AddExport from "./AddExport";
import Footer from "./Footer";
import { useState, useEffect } from "react";
import {
  IPersonaSharedProps,
  Persona,
  PersonaSize,
  PersonaPresence,
} from "office-ui-fabric-react/lib/Persona";

const sortIcon = require("../../../ExternalRef/img/sort.png");
const editIcon = require("../../../ExternalRef/img/EditIcon.png");
const searchIcon = require("../../../ExternalRef/img/searchIcon.png");
const clockIcon = require("../../../ExternalRef/img/clock.png");
const objFilterVal = {
  ID: 0,
  Status: [],
  Priority: null,
  Name: "",
  EngagementType: "",
  EngagementSubType: "",
  UnitName: "",
  CreationDate: null,
  CountryIBVT: "",
  Requestor: "",
  LastModifiedDate: null,
};

let arrCountries = [];
let ArrProjectData = [];
let arrActionData = [];
let arrSpentData = [];

let objSorted = {
  ID: "",
  Status: "",
  Priority: "",
  Name: "",
  EngagementType: "",
  EngagementSubType: "",
  UnitName: "",
  CreationDate: "",
  CountryIBVT: "",
  LatestAction: "",
};

const firstIndex = 0;
let pageSize = 21;
let AdminsArr = [];
let currentUser = "";
let Admin;
let arrPriority = [];
let objSelectedUser = {
  Name: "",
  Email: "",
  JobTitle: "",
};
let arrSelectedRemainig = [];
const App = (props: any) => {
  const [tableData, setTableData] = useState([]);
  const [filterValue, setFilterValues] = useState(objFilterVal);
  const [renderTable, setRenderTable] = useState(false);
  const [countryChoice, setCountryChoice] = useState(arrCountries);
  const [showModal, setShowModal] = useState(false);
  const [showEdit, setShowEdit] = useState({
    flagEdit: false,
    item: null,
    Title: "",
    num: "1",
  });
  const [page, setPage] = useState(1);
  const [data, setData] = useState(tableData.slice(firstIndex, pageSize));
  const [callList, setCallList] = useState(true);
  const [prioLi, setPrioLi] = useState([]);
  const [selectedUserDetails, setSelectedUserDetails] =
    useState(objSelectedUser);
  const [selectedRemainigUsers, setSelectedRemainingUsers] =
    useState(arrSelectedRemainig);
  // const [spentTimeData, setSpentTimeData] = useState([""]);

  // Life Cycle of Onload
  useEffect(() => {
    // get all group users
    props.sp.web.siteGroups
      .getByName("CASAdmin")
      .users()
      .then((res) => {
        AdminsArr = res.map((e) => e.Email.toLowerCase());
      })
      .catch((err) => {
        console.log(err);
      });
    // get current user
    props.sp.web
      .currentUser()
      .then((res) => {
        currentUser = res.Email.toLowerCase();
        Admin = AdminsArr.some((e) => e == currentUser);
      })
      .catch((err) => {
        console.log(err);
      });
    props.sp.web.lists
      .getByTitle("Priorities")
      .items.top(4000)
      .get()
      .then((res) => {
        console.log(res);
        arrPriority = res.map((re) => ({
          Title: re.Title,
          IconUrl: JSON.parse(re.Icon).serverRelativeUrl,
        }));
        setPrioLi(arrPriority);
      });
    // Actions List Call
    props.sp.web.lists
      .getByTitle("Actions")
      .items.top(4000)
      .select("*", "CASRef/ID", "CASAuthor/EMail", "CASAuthor/Title")
      .expand("CASRef", "CASAuthor")
      .orderBy("Created", true)
      .get()
      .then((response) => {
        arrActionData = response.map((res) => {
          return {
            Title: res.Title ? res.Title : "",
            Text: res.CASText ? res.CASText : "",
            Author: res.CASAuthor ? res.CASAuthor.EMail : "",
            AuthorName: res.CASAuthor ? res.CASAuthor.Title : "",
            Ref: res.CASRef ? res.CASRef.ID : 0,
            Created: res.Created ? res.Created : null,
            Modified: res.Modified ? res.Modified : null,
          };
        });
      })
      .then(async () => {
        await props.sp.web.lists
          .getByTitle("Countries")
          .items.top(4000)
          .get()
          .then((cLi) => {
            arrCountries = cLi.map((li) => {
              return li.Title ? li.Title : "";
            });
            setCountryChoice(arrCountries);
          })
          .then(async () => {
            // SpentTimeData
            await props.sp.web.lists
              .getByTitle("Spent time")
              .items.top(4000)
              .get()
              .then((sData) => {
                arrSpentData = sData;
              });
          });
        Admin
          ? await props.sp.web.lists
              .getByTitle("Projects")
              .items.select(
                "*",
                "CASUser/Title",
                "CASUser/ID",
                "CASUser/EMail",
                "CASCountry/Title",
                "CASEngType/Title",
                "CASPriority/Title",
                "CASStatus/Title",
                "CASEngSubType/Title"
              )
              .expand(
                "CASUser",
                "CASCountry",
                "CASEngType",
                "CASPriority",
                "CASStatus",
                "CASEngSubType"
              )
              .top(4000)
              .orderBy("Modified", false)
              .get()
              .then(async (response) => {
                // response = response.filter()
                ArrProjectData = await response.map((item) => {
                  let filteredComments = arrActionData.filter(
                    (aData) => aData.Ref == item.ID
                  );

                  let filteredSpentData = arrSpentData.filter(
                    (aData: any) => aData.CASRefId == item.ID
                  );
                  let requestorDetails = [];
                  requestorDetails = item.CASUser
                    ? item.CASUser.map((user) => ({
                        Name: user.Title,
                        Email: user.EMail,
                        ShowUserDetail: false,
                        ShowRemUsers: false,
                      }))
                    : [];
                  return {
                    ID: item.ID ? item.ID : 0,
                    LatestComment: filteredComments
                      ? filteredComments[filteredComments.length - 1]
                      : [],
                    Status: item.CASStatus.Title ? item.CASStatus.Title : "",
                    Priority: item.CASPriority.Title
                      ? item.CASPriority.Title
                      : "",
                    Name: item.Title ? item.Title : "",
                    EngagementType: item.CASEngType
                      ? item.CASEngType.Title
                      : "",
                    EngagementSubType: item.CASEngSubType
                      ? item.CASEngSubType.Title
                      : "",
                    UnitName: item.CASOrgUnit ? item.CASOrgUnit : "",
                    CreationDate: new Date(item.Created),
                    CountryIBVT: item.CASCountry.Title
                      ? item.CASCountry.Title
                      : "",
                    Requestor: requestorDetails,
                    LastModifiedDate: new Date(item.Modified),
                    LatestActionModified:
                      filteredComments.length > 0
                        ? filteredComments[0].Created
                        : new Date("07/08/1989").toISOString(),
                    ShowRemainingUsers: false,
                    PriorityNo: item.CASPriority.Title
                      ? item.CASPriority.Title.toLowerCase() == "low"
                        ? "1"
                        : item.CASPriority.Title.toLowerCase() == "medium"
                        ? "2"
                        : item.CASPriority.Title.toLowerCase() == "high"
                        ? "3"
                        : ""
                      : "",
                    CrossChargeInfo: item.CASCCI,
                    filteredSpentData: filteredSpentData
                      ? filteredSpentData
                      : [],
                  };
                });
                setRenderTable(true);
              })
          : await props.sp.web.lists
              .getByTitle("Projects")
              .items.select(
                "*",
                "CASUser/Title",
                "CASUser/ID",
                "CASUser/EMail",
                "CASCountry/Title",
                "CASEngType/Title",
                "CASPriority/Title",
                "CASStatus/Title",
                "CASEngSubType/Title"
              )
              .expand(
                "CASUser",
                "CASCountry",
                "CASEngType",
                "CASPriority",
                "CASStatus",
                "CASEngSubType"
              )
              .top(4000)
              .filter(`CASUser/EMail eq '${currentUser}'`)
              .orderBy("Modified", false)
              .get()
              .then(async (response) => {
                // response = response.filter()
                ArrProjectData = await response.map((item) => {
                  let filteredComments = arrActionData.filter(
                    (aData) => aData.Ref == item.ID
                  );
                  let filteredSpentData = arrSpentData.filter(
                    (aData: any) => aData.CASRefId == item.ID
                  );
                  console.log(filteredSpentData);

                  let requestorDetails = [];

                  requestorDetails = item.CASUser
                    ? item.CASUser.map((user) => ({
                        Name: user.Title,
                        Email: user.EMail,
                        ShowUserDetail: false,
                        ShowRemUsers: false,
                      }))
                    : [];
                  return {
                    ID: item.ID ? item.ID : 0,
                    LatestComment: filteredComments
                      ? filteredComments[filteredComments.length - 1]
                      : [],
                    Status: item.CASStatus.Title ? item.CASStatus.Title : "",
                    Priority: item.CASPriority.Title
                      ? item.CASPriority.Title
                      : "",
                    Name: item.Title ? item.Title : "",
                    EngagementType: item.CASEngType
                      ? item.CASEngType.Title
                      : "",
                    EngagementSubType: item.CASEngSubType
                      ? item.CASEngSubType.Title
                      : "",
                    UnitName: item.CASOrgUnit ? item.CASOrgUnit : "",
                    CreationDate: new Date(item.Created),
                    CountryIBVT: item.CASCountry.Title
                      ? item.CASCountry.Title
                      : "",
                    Requestor: requestorDetails,
                    LastModifiedDate: new Date(item.Modified),
                    LatestActionModified:
                      filteredComments.length > 0
                        ? filteredComments[0].Created
                        : new Date("07/08/1989").toISOString(),
                    ShowRemainingUsers: false,
                    PriorityNo: item.CASPriority.Title
                      ? item.CASPriority.Title.toLowerCase() == "low"
                        ? "1"
                        : item.CASPriority.Title.toLowerCase() == "medium"
                        ? "2"
                        : item.CASPriority.Title.toLowerCase() == "high"
                        ? "3"
                        : ""
                      : "",
                    CrossChargeInfo: item.CASCCI,
                    filteredSpentData: filteredSpentData
                      ? filteredSpentData
                      : [],
                  };
                });
                setRenderTable(true);
              });
      })
      .catch((error) => {
        console.log(error);
      });
    setCallList(false);
  }, [callList]);

  useEffect(() => {
    if (
      filterValue.ID == 0 &&
      filterValue.Status.length == 0 &&
      filterValue.Priority == null &&
      filterValue.Name == "" &&
      filterValue.EngagementType == "" &&
      filterValue.EngagementSubType == "" &&
      filterValue.UnitName == "" &&
      filterValue.CreationDate == null &&
      filterValue.CountryIBVT == "" &&
      filterValue.Requestor == "" &&
      filterValue.LastModifiedDate == null
    ) {
      setTableData(ArrProjectData);
      setData(tableData.slice(0, pageSize));
    } else {
      let arrFilteredData = ArrProjectData.filter((fItem) =>
        fItem.Name.toLowerCase().includes(filterValue.Name.toLowerCase())
      );
      arrFilteredData = arrFilteredData.filter((fItem) =>
        filterValue.Priority ? fItem.Priority == filterValue.Priority : true
      );
      arrFilteredData = arrFilteredData.filter((fItem) =>
        filterValue.CountryIBVT
          ? fItem.CountryIBVT == filterValue.CountryIBVT
          : true
      );
      arrFilteredData = arrFilteredData.filter((fItem) =>
        fItem.UnitName.toLowerCase().includes(
          filterValue.UnitName.toLowerCase()
        )
      );
      arrFilteredData = arrFilteredData.filter((fItem) =>
        filterValue.Requestor
          ? fItem.Requestor.map((req) => req.Email).some(
              (userMail) => userMail == filterValue.Requestor
            )
          : true
      );
      arrFilteredData = arrFilteredData.filter((fItem) =>
        filterValue.EngagementType
          ? filterValue.EngagementType == fItem.EngagementType
          : true
      );
      arrFilteredData = arrFilteredData.filter((fItem) =>
        filterValue.EngagementSubType
          ? filterValue.EngagementSubType == fItem.EngagementSubType
          : true
      );
      // arrFilteredData = arrFilteredData.filter((fItem) =>
      //   filterValue.Status ? filterValue.Status == fItem.Status : true
      // );
      arrFilteredData = arrFilteredData.filter((fItem) =>
        filterValue.Status.length > 0 && fItem.Status
          ? filterValue.Status["includes"](fItem.Status)
          : filterValue.Status.length > 0
          ? false
          : true
      );
      arrFilteredData = arrFilteredData.filter((fItem) =>
        filterValue.ID != 0 ? filterValue.ID == fItem.ID : true
      );
      arrFilteredData = arrFilteredData.filter((fItem) =>
        filterValue.CreationDate && filterValue.CreationDate != "Invalid Date"
          ? new Date(filterValue.CreationDate).toLocaleDateString() ==
            new Date(fItem.CreationDate).toLocaleDateString()
          : true
      );
      arrFilteredData = arrFilteredData.filter((fItem) =>
        filterValue.LastModifiedDate &&
        filterValue.LastModifiedDate != "Invalid Date"
          ? new Date(filterValue.LastModifiedDate).toLocaleDateString() ==
            new Date(fItem.LastModifiedDate).toLocaleDateString()
          : true
      );
      setTableData(arrFilteredData);
      setData(tableData.slice(0, pageSize));
    }
    setRenderTable(false);
  }, [renderTable]);

  // get on Filter Items
  const getFilterItems = (data) => {
    setFilterValues({ ...data });
    setRenderTable(true);
  };

  const handleChange = (event, value) => {
    setPage(value);
    setData(
      tableData.slice(firstIndex + pageSize * (value - 1), pageSize * value)
    );
  };

  const renderList = () => {
    setCallList(true);
  };

  // const latestEdit = (data, text) => {
  //   setShowModal(true);
  //   setlatest({ Id: data, text: text });
  // };
  const getModalResponse = (res) => {
    setShowModal(res);
  };

  const Edit = (key, IdValue, data, num) => {
    setShowModal(true);
    setShowEdit({ flagEdit: key, item: IdValue, Title: data, num: num });
  };
  const getSelectedUserDetails = (userMail) => {
    props.sp.profiles
      .getPropertiesFor(`i:0#.f|membership|${userMail}`)
      .then((res) => {
        (objSelectedUser = {
          Name: res.DisplayName,
          Email: userMail,
          JobTitle: res.UserProfileProperties.filter(
            (uProfile) => uProfile.Key == "SPS-JobTitle"
          )[0].Value,
        }),
          console.log(objSelectedUser);
        setSelectedUserDetails({ ...objSelectedUser });
      })
      .catch((err) => console.log(err));
  };
  const getRemainingUserDetails = (arrUsers) => {
    arrSelectedRemainig = [];
    setSelectedRemainingUsers([]);
    arrUsers.forEach((user, i) => {
      if (i > 2) {
        props.sp.profiles
          .getPropertiesFor(`i:0#.f|membership|${user.Email}`)
          .then((res) => {
            arrSelectedRemainig.push({
              Name: res.DisplayName,
              Email: user.Email,
              JobTitle: res.UserProfileProperties.filter(
                (uProfile) => uProfile.Key == "SPS-JobTitle"
              )[0].Value,
            });

            if (i == arrUsers.length - 1) {
              console.log(arrSelectedRemainig);
              setTimeout(() => {
                setSelectedRemainingUsers([]);
                setSelectedRemainingUsers(arrSelectedRemainig);
              }, 100);
            }
            // return {
            //   Name: res.DisplayName,
            //   Email: user.Email,
            //   JobTitle: res.UserProfileProperties.filter(
            //     (uProfile) => uProfile.Key == "SPS-JobTitle"
            //   )[0].Value,
            // };
          })
          .catch((err) => console.log(err));
      }
    });
    // return arrSelectedRemainig;
    // setSelectedRemainingUsers(arrSelectedRemainig);
  };
  return (
    <>
      {/* <Header sp={props.sp} /> */}
      <div style={{ position: "relative" }}>
        {showModal && (
          <Panel
            Panel={getModalResponse}
            Edit={showEdit}
            context={props.context}
            sp={props.sp}
            renderProject={renderList}
            Admin={Admin}
          />
        )}
        <AddExport
          Panel={getModalResponse}
          Edit={Edit}
          Admin={Admin}
          exportData={tableData}
        />
        <TopFilter
          context={props.context}
          sp={props.sp}
          filterdata={getFilterItems}
          data={ArrProjectData}
        />
        <TableContainer
          style={{
            width: "98%",
            margin: "auto",
            padding: "2rem",
            overflowX: "auto",
          }}
        >
          <Table className={classes.projectTable}>
            <TableHead>
              <TableRow>
                <TableCell
                  onClick={() => {
                    objSorted.ID == "ascending" || objSorted.ID == ""
                      ? (objSorted = {
                          ID: "descending",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        })
                      : (objSorted = {
                          ID: "ascending",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        });
                    setTableData([
                      ...tableData.sort((a, b) =>
                        objSorted.ID == "ascending" || objSorted.ID == ""
                          ? b.ID - a.ID
                          : a.ID - b.ID
                      ),
                    ]);
                    setData(tableData.slice(0, pageSize));
                    // setRenderTable(true);
                  }}
                >
                  <div style={{ color: "#7d7d7d", display: "flex" }}>
                    ID{" "}
                    <div>
                      <img height="10" width="10" src={`${sortIcon}`} />
                    </div>
                  </div>
                </TableCell>
                <TableCell
                  onClick={() => {
                    objSorted.Status == "ascending" || objSorted.Status == ""
                      ? (objSorted = {
                          ID: "",
                          Status: "descending",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        })
                      : (objSorted = {
                          ID: "",
                          Status: "ascending",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        });
                    setTableData([
                      ...tableData.sort((a, b) =>
                        objSorted.Status == "ascending" ||
                        objSorted.Status == ""
                          ? b.Status.toLowerCase().localeCompare(
                              a.Status.toLowerCase()
                            )
                          : a.Status.toLowerCase().localeCompare(
                              b.Status.toLowerCase()
                            )
                      ),
                    ]);
                    setData(tableData.slice(0, pageSize));
                  }}
                >
                  <div style={{ color: "#7d7d7d", display: "flex" }}>
                    Status{" "}
                    <div>
                      <img height="10" width="10" src={`${sortIcon}`} />
                    </div>
                  </div>
                </TableCell>
                <TableCell
                  onClick={() => {
                    objSorted.Priority == "ascending" ||
                    objSorted.Priority == ""
                      ? (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "descending",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        })
                      : (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "ascending",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        });
                    setTableData([
                      ...tableData.sort((a, b) =>
                        objSorted.Priority == "ascending" ||
                        objSorted.Priority == ""
                          ? b.PriorityNo - a.PriorityNo
                          : a.PriorityNo - b.PriorityNo
                      ),
                    ]);
                    setData([...tableData.slice(0, pageSize)]);
                  }}
                >
                  <div style={{ color: "#7d7d7d", display: "flex" }}>
                    Priority{" "}
                    <div>
                      <img height="10" width="10" src={`${sortIcon}`} />
                    </div>
                  </div>
                </TableCell>
                <TableCell
                  onClick={() => {
                    objSorted.Name == "ascending" || objSorted.Name == ""
                      ? (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "descending",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        })
                      : (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "ascending",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        });
                    setTableData([
                      ...tableData.sort((a, b) =>
                        objSorted.Name == "ascending" || objSorted.Name == ""
                          ? b.Name.toLowerCase().localeCompare(
                              a.Name.toLowerCase()
                            )
                          : a.Name.toLowerCase().localeCompare(
                              b.Name.toLowerCase()
                            )
                      ),
                    ]);
                    setData(tableData.slice(0, pageSize));
                  }}
                >
                  <div style={{ color: "#7d7d7d", display: "flex" }}>
                    Name{" "}
                    <div>
                      <img height="10" width="10" src={`${sortIcon}`} />
                    </div>
                  </div>
                </TableCell>
                <TableCell
                  onClick={() => {
                    objSorted.EngagementType == "ascending" ||
                    objSorted.EngagementType == ""
                      ? (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "descending",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        })
                      : (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "ascending",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        });
                    setTableData([
                      ...tableData.sort((a, b) =>
                        b.EngagementType && a.EngagementType
                          ? objSorted.EngagementType == "ascending" ||
                            objSorted.EngagementType == ""
                            ? b.EngagementType.toLowerCase().localeCompare(
                                a.EngagementType.toLowerCase()
                              )
                            : a.EngagementType.toLowerCase().localeCompare(
                                b.EngagementType.toLowerCase()
                              )
                          : ""
                      ),
                    ]);
                    setData(tableData.slice(0, pageSize));
                  }}
                >
                  <div style={{ color: "#7d7d7d", display: "flex" }}>
                    Engagement type{" "}
                    <div>
                      <img height="10" width="10" src={`${sortIcon}`} />
                    </div>
                  </div>
                </TableCell>
                <TableCell
                  onClick={() => {
                    objSorted.EngagementSubType == "ascending" ||
                    objSorted.EngagementSubType == ""
                      ? (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "descending",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        })
                      : (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "ascending",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        });
                    setTableData([
                      ...tableData.sort((a, b) =>
                        b.EngagementSubType && a.EngagementSubType
                          ? objSorted.EngagementSubType == "ascending" ||
                            objSorted.EngagementSubType == ""
                            ? b.EngagementSubType.toLowerCase().localeCompare(
                                a.EngagementSubType.toLowerCase()
                              )
                            : a.EngagementSubType.toLowerCase().localeCompare(
                                b.EngagementSubType.toLowerCase()
                              )
                          : ""
                      ),
                    ]);
                    setData(tableData.slice(0, pageSize));
                  }}
                >
                  <div style={{ color: "#7d7d7d", display: "flex" }}>
                    Engagement subtype{" "}
                    <div>
                      <img height="10" width="10" src={`${sortIcon}`} />
                    </div>
                  </div>
                </TableCell>
                <TableCell
                  onClick={() => {
                    objSorted.UnitName == "ascending" ||
                    objSorted.UnitName == ""
                      ? (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "descending",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        })
                      : (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "ascending",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "",
                        });
                    setTableData([
                      ...tableData.sort((a, b) =>
                        b.UnitName && a.UnitName
                          ? objSorted.UnitName == "ascending" ||
                            objSorted.UnitName == ""
                            ? b.UnitName.toLowerCase().localeCompare(
                                a.UnitName.toLowerCase()
                              )
                            : a.UnitName.toLowerCase().localeCompare(
                                b.UnitName.toLowerCase()
                              )
                          : ""
                      ),
                    ]);
                    setData(tableData.slice(0, pageSize));
                  }}
                >
                  <div style={{ color: "#7d7d7d", display: "flex" }}>
                    Organization Unit{" "}
                    <div>
                      <img height="10" width="10" src={`${sortIcon}`} />
                    </div>
                  </div>
                </TableCell>
                <TableCell
                  onClick={() => {
                    objSorted.CreationDate == "ascending" ||
                    objSorted.CreationDate == ""
                      ? (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "descending",
                          CountryIBVT: "",
                          LatestAction: "",
                        })
                      : (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "ascending",
                          CountryIBVT: "",
                          LatestAction: "",
                        });
                    setTableData([
                      ...tableData.sort((a, b) =>
                        objSorted.CreationDate == "ascending" ||
                        objSorted.CreationDate == ""
                          ? Date.parse(a.CreationDate) -
                            Date.parse(b.CreationDate)
                          : Date.parse(b.CreationDate) -
                            Date.parse(a.CreationDate)
                      ),
                    ]);
                    setData(tableData.slice(0, pageSize));
                  }}
                >
                  <div style={{ color: "#7d7d7d", display: "flex" }}>
                    Creation Date{" "}
                    <div>
                      <img height="10" width="10" src={`${sortIcon}`} />
                    </div>
                  </div>
                </TableCell>
                <TableCell
                  onClick={() => {
                    objSorted.CountryIBVT == "ascending" ||
                    objSorted.CountryIBVT == ""
                      ? (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "descending",
                          LatestAction: "",
                        })
                      : (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "ascending",
                          LatestAction: "",
                        });
                    setTableData([
                      ...tableData.sort((a, b) =>
                        b.CountryIBVT && a.CountryIBVT
                          ? objSorted.CountryIBVT == "ascending" ||
                            objSorted.CountryIBVT == ""
                            ? b.CountryIBVT.toLowerCase().localeCompare(
                                a.CountryIBVT.toLowerCase()
                              )
                            : a.CountryIBVT.toLowerCase().localeCompare(
                                b.CountryIBVT.toLowerCase()
                              )
                          : ""
                      ),
                    ]);
                    setData(tableData.slice(0, pageSize));
                  }}
                >
                  <div style={{ color: "#7d7d7d", display: "flex" }}>
                    Country/IBVT{" "}
                    <div>
                      <img height="10" width="10" src={`${sortIcon}`} />
                    </div>
                  </div>
                </TableCell>
                <TableCell>
                  <div style={{ color: "#7d7d7d", display: "flex" }}>
                    Requestor{" "}
                  </div>
                </TableCell>
                <TableCell
                  onClick={() => {
                    objSorted.LatestAction == "ascending" ||
                    objSorted.LatestAction == ""
                      ? (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "descending",
                        })
                      : (objSorted = {
                          ID: "",
                          Status: "",
                          Priority: "",
                          Name: "",
                          EngagementType: "",
                          EngagementSubType: "",
                          UnitName: "",
                          CreationDate: "",
                          CountryIBVT: "",
                          LatestAction: "ascending",
                        });

                    setTableData([
                      ...tableData.sort((a, b) =>
                        //  (a.LatestComment && b.LatestComment) &&
                        objSorted.LatestAction == "ascending" ||
                        objSorted.LatestAction == ""
                          ? Date.parse(a.LatestActionModified) -
                            Date.parse(b.LatestActionModified)
                          : Date.parse(b.LatestActionModified) -
                            Date.parse(a.LatestActionModified)
                      ),
                    ]);
                    setData(tableData.slice(0, pageSize));
                  }}
                >
                  <div style={{ color: "#7d7d7d", display: "flex" }}>
                    Latest action{" "}
                    <div>
                      <img height="10" width="10" src={`${sortIcon}`} />
                    </div>
                  </div>
                </TableCell>
                <TableCell></TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              <>
                {data.length > 0 &&
                  data.map((row, i) => {
                    return (
                      <TableRow
                        key={i}
                        onMouseEnter={() => {
                          data.forEach((dT) => {
                            dT.ShowRemainingUsers = false;
                          });
                          data.forEach((dT) => {
                            dT.Requestor.forEach((person) => {
                              person.ShowUserDetail = false;
                            });
                          });
                          setData([...data]);
                        }}
                      >
                        <TableCell style={{ width: "20px" }}>
                          <div className={classes.TableID}> {row.ID}</div>
                        </TableCell>
                        <TableCell style={{ width: "100px" }}>
                          <div className={classes.statusSection}>
                            {row.Status && (
                              <div
                                className={classes.statusCircle}
                                style={{
                                  background:
                                    row.Status.toLowerCase() == "in progress"
                                      ? "#359942"
                                      : row.Status.toLowerCase() ==
                                        "waiting for feedback"
                                      ? "#f5944e"
                                      : row.Status.toLowerCase() == "lead"
                                      ? "#f24998"
                                      : row.Status.toLowerCase() == "parked"
                                      ? "#999999"
                                      : row.Status.toLowerCase() == "closed"
                                      ? "#1c75bc"
                                      : row.Status.toLowerCase() == "canceled"
                                      ? "#7e2e7a"
                                      : "#000",
                                }}
                              ></div>
                            )}
                            <div
                              className={classes.statusName}
                              style={{
                                color:
                                  row.Status.toLowerCase() == "in progress"
                                    ? "#359942"
                                    : row.Status.toLowerCase() ==
                                      "waiting for feedback"
                                    ? "#f5944e"
                                    : row.Status.toLowerCase() == "lead"
                                    ? "#f24998"
                                    : row.Status.toLowerCase() == "parked"
                                    ? "#999999"
                                    : row.Status.toLowerCase() == "closed"
                                    ? "#1c75bc"
                                    : row.Status.toLowerCase() == "canceled"
                                    ? "#7e2e7a"
                                    : "#000",
                              }}
                            >
                              {row.Status}
                            </div>
                          </div>
                        </TableCell>
                        <TableCell
                          style={{
                            width: 50,
                          }}
                        >
                          {/* {row.Priority} */}
                          <div className={classes.prioritySection}>
                            {row.Priority != "" ? (
                              <img
                                src={
                                  prioLi.filter(
                                    (li) => li.Title == row.Priority
                                  )[0].IconUrl
                                }
                                width={30}
                                height={30}
                              />
                            ) : (
                              ""
                            )}
                          </div>
                        </TableCell>
                        <TableCell style={{ fontSize: 20, width: 150 }}>
                          <div className={`${classes.bold} ${classes.PName}`}>
                            {row.Name}
                          </div>
                        </TableCell>
                        <TableCell
                          style={{
                            fontSize: 14,
                            fontWeight: 600,
                            color: "#707070",
                            width: 150,
                          }}
                        >
                          <div className={classes.bold}>
                            {row.EngagementType}
                          </div>
                        </TableCell>
                        <TableCell
                          style={{
                            fontSize: 14,
                            fontWeight: 600,
                            color: "#707070",
                            width: 160,
                          }}
                        >
                          <div className={classes.bold}>
                            {row.EngagementSubType}
                          </div>
                        </TableCell>
                        <TableCell
                          style={{
                            fontSize: 14,
                            fontWeight: 600,
                            color: "#707070",
                            width: 160,
                          }}
                        >
                          <div className={classes.bold}>{row.UnitName}</div>
                        </TableCell>
                        <TableCell
                          style={{
                            fontSize: 14,
                            color: "#707070",
                            width: 160,
                          }}
                        >
                          {/* {new Date(row.CreationDate).getDate() +
                          "/" +
                          (new Date(row.CreationDate).getMonth() + 1) +
                          "/" +
                          new Date(row.CreationDate).getFullYear()} */}
                          {/* {DateFormatter(row.CreationDate)} */}
                          <div className={classes.normal}>
                            {`${
                              +new Date(row.CreationDate)
                                .toLocaleDateString()
                                .split("/")[1] < 10
                                ? "0" +
                                  new Date(row.CreationDate)
                                    .toLocaleDateString()
                                    .split("/")[1]
                                : new Date(row.CreationDate)
                                    .toLocaleDateString()
                                    .split("/")[1]
                            }/${
                              +new Date(row.CreationDate)
                                .toLocaleDateString()
                                .split("/")[0] < 10
                                ? "0" +
                                  new Date(row.CreationDate)
                                    .toLocaleDateString()
                                    .split("/")[0]
                                : new Date(row.CreationDate)
                                    .toLocaleDateString()
                                    .split("/")[0]
                            }/${new Date(row.CreationDate)
                              .toLocaleDateString()
                              .split("/")[2]
                              .toString()
                              .substr(-2)}`}
                            {/* {new Date(row.CreationDate).toLocaleDateString()} */}
                          </div>
                        </TableCell>
                        <TableCell
                          style={{
                            fontSize: 14,
                            color: "#707070",
                            width: 80,
                          }}
                          onMouseEnter={() => {
                            data.forEach((dT) => {
                              dT.ShowRemainingUsers = false;
                            });
                            data.forEach((dT) => {
                              dT.Requestor.forEach((person) => {
                                person.ShowUserDetail = false;
                              });
                            });
                            setData([...data]);
                          }}
                        >
                          <div className={classes.normal}>
                            {row.CountryIBVT}
                          </div>
                        </TableCell>
                        <TableCell
                          style={{
                            width: 60,
                            padding: 0,
                          }}
                          onClick={() => {
                            data.forEach((dT) => {
                              dT.ShowRemainingUsers = false;
                            });
                            data.forEach((dT) => {
                              dT.Requestor.forEach((person) => {
                                person.ShowUserDetail = false;
                              });
                            });
                          }}
                        >
                          <div className={classes.PeopleIcons}>
                            {row.Requestor.map((peopleIcon, i) =>
                              i < 3 ? (
                                <div className={classes.RequestorSection}>
                                  {" "}
                                  <Persona
                                    styles={{
                                      root: {
                                        width: 48,
                                        position: "relative",
                                        zIndex: 3 - i,
                                        left: i * -20,
                                        cursor: "pointer",
                                      },
                                    }}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      peopleIcon.Email
                                    }
                                    onMouseLeave={() => {
                                      data.forEach((dT) => {
                                        dT.ShowRemainingUsers = false;
                                      });
                                      data.forEach((dT) => {
                                        dT.Requestor.forEach((person) => {
                                          person.ShowUserDetail = false;
                                        });
                                      });
                                    }}
                                    onMouseEnter={() => {
                                      getSelectedUserDetails(peopleIcon.Email);
                                      data.forEach((dT) => {
                                        dT.ShowRemainingUsers = false;
                                      });
                                      data.forEach((dT) => {
                                        dT.Requestor.forEach((person) => {
                                          person.ShowUserDetail = false;
                                        });
                                      });
                                      data
                                        .filter((tD) => tD.ID == row.ID)[0]
                                        .Requestor.filter(
                                          (req) => req.Email == peopleIcon.Email
                                        )[0].ShowUserDetail = true;
                                      setData([
                                        ...data.slice(firstIndex, pageSize),
                                      ]);
                                    }}
                                  />
                                  {peopleIcon.ShowUserDetail && (
                                    <div className={classes.userDetails}>
                                      <div
                                        className={classes.userDetailsHeader}
                                      >
                                        <div className={classes.userDetailsDP}>
                                          <Persona
                                            styles={{
                                              root: {
                                                width: 72,
                                                margin: "auto",
                                              },
                                            }}
                                            imageUrl={
                                              "/_layouts/15/userphoto.aspx?size=S&username=" +
                                              peopleIcon.Email
                                            }
                                            size={PersonaSize.size72}
                                          />
                                        </div>
                                        <div className={classes.userContent}>
                                          <div className={classes.userName}>
                                            {selectedUserDetails.Name}
                                          </div>
                                          <div className={classes.userjobTitle}>
                                            {selectedUserDetails.JobTitle}
                                          </div>
                                          <div className={classes.userEmail}>
                                            <a
                                              href={`mailto:${selectedUserDetails.Email}`}
                                            >
                                              {selectedUserDetails.Email}
                                            </a>
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  )}
                                </div>
                              ) : (
                                ""
                              )
                            )}
                            {row.Requestor.length > 3 ? (
                              <div className={classes.RequestorSection}>
                                <div
                                  onMouseLeave={() => {
                                    data.forEach((dT) => {
                                      dT.ShowRemainingUsers = false;
                                    });
                                    data.forEach((dT) => {
                                      dT.Requestor.forEach((person) => {
                                        person.ShowUserDetail = false;
                                      });
                                    });
                                  }}
                                  onMouseEnter={() => {
                                    setSelectedRemainingUsers([]);
                                    getRemainingUserDetails(row.Requestor);

                                    data.forEach((dT) => {
                                      dT.ShowRemainingUsers = false;
                                    });
                                    data.forEach((dT) => {
                                      dT.Requestor.forEach((person) => {
                                        person.ShowUserDetail = false;
                                      });
                                    });
                                    data.filter(
                                      (dT) => dT.ID == row.ID
                                    )[0].ShowRemainingUsers = true;
                                    setData([
                                      ...data.slice(firstIndex, pageSize),
                                    ]);
                                  }}
                                  style={{
                                    width: 48,
                                    height: 48,
                                    position: "relative",
                                    left: -53,
                                    borderRadius: "50%",
                                    fontSize: "1.5rem",
                                    fontWeight: "bold",
                                    color: "#fff",
                                    display: "flex",
                                    justifyContent: "center",
                                    alignItems: "center",
                                    backgroundColor: "#00e8d1",
                                    cursor: "pointer",
                                  }}
                                >
                                  +{row.Requestor.length - 3}
                                </div>

                                {row.ShowRemainingUsers && (
                                  <div className={classes.userDetails}>
                                    {selectedRemainigUsers.map((remainUser) => (
                                      <div className={classes.singleUser}>
                                        <div>
                                          <Persona
                                            styles={{
                                              root: {
                                                width: 72,
                                                margin: "auto",
                                              },
                                            }}
                                            imageUrl={
                                              "/_layouts/15/userphoto.aspx?size=S&username=" +
                                              remainUser.Email
                                            }
                                            size={PersonaSize.size48}
                                          />
                                        </div>
                                        <div>
                                          <div className={classes.singleName}>
                                            {remainUser.Name}
                                          </div>
                                          <div
                                            className={classes.singleJobTitle}
                                          >
                                            {remainUser.JobTitle}
                                          </div>
                                          <div className={classes.singleEmail}>
                                            <a>{remainUser.Email}</a>
                                          </div>
                                        </div>
                                      </div>
                                    ))}
                                  </div>
                                )}
                              </div>
                            ) : (
                              ""
                            )}
                          </div>
                        </TableCell>
                        <TableCell
                          style={{
                            fontSize: 14,
                            color: "#707070",
                            width: 200,
                          }}
                          onMouseEnter={() => {
                            data.forEach((dT) => {
                              dT.ShowRemainingUsers = false;
                            });
                            data.forEach((dT) => {
                              dT.Requestor.forEach((person) => {
                                person.ShowUserDetail = false;
                              });
                            });
                            setData([...data]);
                          }}
                        >
                          {row.LatestComment && (
                            <div className={classes.LatestActionInLi}>
                              <div className={classes.LAHeader}>
                                <div className={classes.LADp}>
                                  {row.LatestComment && (
                                    <Persona
                                      size={PersonaSize.size32}
                                      styles={{
                                        root: {
                                          marginRight: "0.5rem",
                                        },
                                      }}
                                      imageUrl={
                                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                                        // peopleIcon.EMail
                                        row.LatestComment.Author
                                      }
                                    />
                                  )}
                                </div>
                                <div className={classes.LAUserName}>
                                  {row.LatestComment &&
                                    row.LatestComment.AuthorName}
                                </div>
                                <div className={classes.LAPostedTime}>
                                  {row.LatestComment && (
                                    <div
                                      style={{
                                        display: "flex",
                                        alignItems: "center",
                                      }}
                                    >
                                      <img
                                        src={`${clockIcon}`}
                                        width={12}
                                        height={12}
                                        style={{ marginRight: "0.3rem" }}
                                      />
                                      {+new Date(row.LatestComment.Created)
                                        .toLocaleDateString()
                                        .split("/")[1] < 10
                                        ? "0" +
                                          new Date(row.LatestComment.Created)
                                            .toLocaleDateString()
                                            .split("/")[1]
                                        : new Date(row.LatestComment.Created)
                                            .toLocaleDateString()
                                            .split("/")[1]}
                                      /
                                      {+new Date(row.LatestComment.Created)
                                        .toLocaleDateString()
                                        .split("/")[0] < 10
                                        ? "0" +
                                          new Date(row.LatestComment.Created)
                                            .toLocaleDateString()
                                            .split("/")[0]
                                        : new Date(row.LatestComment.Created)
                                            .toLocaleDateString()
                                            .split("/")[0]}
                                      /
                                      {new Date(row.LatestComment.Created)
                                        .toLocaleDateString()
                                        .split("/")[2]
                                        .toString()
                                        .substr(-2)}
                                    </div>
                                  )}
                                </div>
                              </div>
                              <div
                                className={`${classes.LABody} ${classes.normal}`}
                              >
                                {row.LatestComment &&
                                row.LatestComment.Text.length > 145
                                  ? `${row.LatestComment.Text.substr(
                                      0,
                                      144
                                    )} . . .`
                                  : row.LatestComment && row.LatestComment.Text}
                              </div>
                              <div>
                                <div
                                  style={{
                                    width: "2rem",
                                    height: "2rem",
                                    borderRadius: "50%",
                                    marginLeft: "auto",
                                  }}
                                >
                                  <img
                                    src={`${searchIcon}`}
                                    width={26}
                                    height={26}
                                    style={{ transform: "rotate(90deg)" }}
                                    onClick={() =>
                                      Edit(true, row.ID, row.Name, "2")
                                    }
                                  />
                                </div>
                              </div>
                            </div>
                          )}
                        </TableCell>
                        <TableCell style={{ width: 100 }}>
                          <div
                            style={{
                              width: "3rem",
                              height: "3rem",
                              borderRadius: "50%",
                              margin: "auto",
                            }}
                          >
                            <img
                              style={{ cursor: "pointer" }}
                              src={`${editIcon}`}
                              width={39}
                              height={39}
                              onClick={() => {
                                Edit(true, row.ID, row.Name, "1");
                              }}
                            />
                          </div>
                        </TableCell>
                      </TableRow>
                    );
                  })}
              </>
            </TableBody>
          </Table>
          {data.length == 0 && (
            <div className={classes.noDataFound}>No data found</div>
          )}
        </TableContainer>
        <Pagination
          count={Math.ceil(tableData.length / pageSize)}
          page={page}
          onChange={handleChange}
          className={classes.pagination}
        />
      </div>
      <Footer />
    </>
  );
};
export default App;
