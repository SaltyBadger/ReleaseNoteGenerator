/// <reference path="../node_modules/vss-web-extension-sdk/typings/index.d.ts" />

//imports.
import { BaseControl } from "VSS/Controls";
import { Combo, IComboOptions, ComboDateBehavior } from "VSS/Controls/Combos";
import {
    GitRepository, GitBranchStats, GitQueryBranchStatsCriteria, GitCommit, GitCommitRef, GitQueryCommitsCriteria,
    GitRef, GitVersionDescriptor, GitVersionType, GitPullRequest, PullRequestStatus, GitPullRequestSearchCriteria
} from "TFS/VersionControl/Contracts";
import { getClient as getGitClient } from "TFS/VersionControl/GitRestClient";
import { getClient as getWIClient } from "TFS/WorkItemTracking/RestClient";
import { Grid, IGridOptions } from "VSS/Controls/Grids";
import { IWaitControlOptions, WaitControl } from "VSS/Controls/StatusIndicator";
import TreeView = require("VSS/Controls/TreeView");
import { ResourceRef } from "VSS/WebApi/Contracts";
import { WorkItem } from "TFS/WorkItemTracking/Contracts";
import wIContracts = require("TFS/WorkItemTracking/Contracts")


//params and controlls.
var container = $(".input-controls");
var waitControlOptions: IWaitControlOptions = {
    cancellable: true
};
var newHeight = +($(window).height()) / 2.3;
var waitControler = BaseControl.create(WaitControl, container, waitControlOptions);
var releaseWaitControler = BaseControl.create(WaitControl, container, waitControlOptions);
var resultArr = [];
var reqResultArr = [];
var repoControl = <Combo>BaseControl.createIn(Combo, $(".repoPicker"), <IComboOptions>{ allowEdit: false });
var branchControl = <Combo>BaseControl.createIn(Combo, $(".branchPicker"), <IComboOptions>{ allowEdit: false });
var fromTagControl = <Combo>BaseControl.createIn(Combo, $(".fromTagPicker"), <IComboOptions>{ allowEdit: false });
var toTagControl = <Combo>BaseControl.createIn(Combo, $(".toTagPicker"), <IComboOptions>{ allowEdit: false });
var startDateControl = <Combo>BaseControl.createIn(Combo, $(".fromDate"), <IComboOptions>{ type: 'date-time', value: '1/1/1970' });
var endDateControl = <Combo>BaseControl.createIn(Combo, $(".toDate"), <IComboOptions>{ type: 'date-time', value: formatDate(new Date()) });
var reqGridControl = <Grid>BaseControl.createIn(Grid, $('.reqResults'), {
    height: newHeight, // Explicit height is required for a Grid control
    columns: [
        { text: "Repository", index: "repo", width: 150 },
        { text: "Target Branch", index: "branch", width: 150 },
        { text: "Date", index: "date", width: 150 },
        { text: "Pull Request ID", index: "id", width: 150 },
        { text: "Title", index: "title", width: 300 },
        { text: "Description", index: "desc", width: 600 }
    ],
    source: reqResultArr
});
var gridControl = <Grid>BaseControl.createIn(Grid, $('.results'), {
    height: newHeight, // Explicit height is required for a Grid control
    columns: [
        { text: "Repository", index: "repo", width: 150 },
        { text: "Branch", index: "branch", width: 150 },
        { text: "Date", index: "date", width: 150 },
        { text: "Commit ID", index: "id", width: 200 },
        { text: "Commit", index: "commit", width: 900 }
    ],
    source: resultArr
});
var treeControl;

//main events.
$(document).ready(function () {
    // $(window).bind('resize', set_body_height);
    // set_body_height();
    waitControler.startWait();
    getRepo()
    //waitControler.endWait();
    repoControl._bind('change', () => {
        waitControler.startWait();
        branchControl.setInputText('');
        fromTagControl.setInputText('');
        toTagControl.setInputText('');
        let p = []
        p.push(getBranches());
        p.push(getCommits());
        p.push(getTags());
        Promise.all(p).then(() => { waitControler.endWait() })

    });
    branchControl._bind('change', () => {
        waitControler.startWait();
        getCommits();

    });
    fromTagControl._bind('change', () => {
        waitControler.startWait();
        getCommits();

    });
    toTagControl._bind('change', () => {
        waitControler.startWait();
        getCommits();

    });
    startDateControl._bind('change', () => {
        waitControler.startWait();
        getCommits();

    });
    endDateControl._bind('change', () => {
        waitControler.startWait();
        getCommits();

    });
    let requestID;
    reqGridControl._bind('click', (e) => {
        if (e.target.childElementCount == 0) {
            requestID = e.target.parentElement.childNodes[3].innerText;
            getPullData(requestID);
        }
    });
    let counter = 0;
    gridControl._canvas.scroll(() => {
        counter++;
        if (counter > 5) {
            getPullData(requestID);
            counter = 0;
        }
    });
    $('.refreshBtn').click(() => {
        waitControler.startWait();
        if (repoControl.getText() == '') {
            waitControler.endWait();
        }
        else {
            getCommits();
        }
        if ($('.releaseTable').is(":visible")) {
            treeControl.dispose();
            releaseWaitControler.startWait();
            createTree();
        }
    });
    $('.releaseBtn').click(() => {
        releaseWaitControler.startWait();
        createTree();
    });
    $('.markDownBtn').click(() => {
        releaseWaitControler.startWait();
        createMarkDown();
    });
    $('.resetBtn').click(() => {
        repoControl.setInputText('');
        branchControl.setInputText('');
        fromTagControl.setInputText('');
        toTagControl.setInputText('');
        branchControl.setSource(new Array);
        fromTagControl.setSource(new Array);
        toTagControl.setSource(new Array);
        gridControl.setDataSource(new Array);
        reqGridControl.setDataSource(new Array);
        if (treeControl)
            treeControl.dispose();
    });
});

async function getRepo() {
    let repositories: GitRepository[];
    var repos = await getGitClient().getRepositories(VSS.getWebContext().project.id)
    repositories = repos.sort((a, b) => a.name.localeCompare(b.name));
    repoControl.setSource(repositories.map((r) => r.name));
    waitControler.endWait();
};

async function getBranches() {
    let gitBranches: GitBranchStats[];
    let branches = await getGitClient().getBranches(repoControl.getText(), VSS.getWebContext().project.id)
    gitBranches = branches.sort((a, b) => a.name.localeCompare(b.name));
    branchControl.setSource(gitBranches.map((r) => r.name));
};
async function getTags() {
    let gitBranches: GitBranchStats[];
    let gitRefs = [];
    var tags = await getGitClient().getRefs(repoControl.getText(), VSS.getWebContext().project.id, null, false, false, false, false, true);
    tags.forEach(ref => {
        if (ref.name.split('/')[1] == 'tags') {
            gitRefs.push(ref.name.split('/')[2]);
        }
    });

    fromTagControl.setSource(gitRefs.map((r) => r));
    toTagControl.setSource(gitRefs.map((r) => r));
    return gitRefs;
};

//get commits for its grid controll, and return the current pullrequests associated with the commits.
async function getCommits() {
    let commitsFilter;
    let commits: GitCommitRef[];
    var pullreqInGrid: GitPullRequest[] = [];
    let branchTagLabel = branchControl.getText();
    if (fromTagControl.getText() != '' && toTagControl.getText() != '' && branchControl.getText() != '') {
        commits = await getBatchCommits(fromTagControl.getText(), toTagControl.getText())
    }
    else if (branchControl.getText() != '') {
        commitsFilter = <GitQueryCommitsCriteria>{
            itemVersion: <GitVersionDescriptor>{ version: branchControl.getText() },
            fromDate: startDateControl.getText(), toDate: nextDay(endDateControl.getText())
        };
        branchTagLabel = branchControl.getText();
        commits = await getGitClient().getCommits(repoControl.getText(), commitsFilter, VSS.getWebContext().project.id);
    }
    else {
        commitsFilter = <GitQueryCommitsCriteria>{ fromDate: startDateControl.getText(), toDate: nextDay(endDateControl.getText()) };
        commits = await getGitClient().getCommits(repoControl.getText(), commitsFilter, VSS.getWebContext().project.id);
    };
    let targetBranch = '';
    if (branchControl.getText() != '')
        targetBranch = 'refs/heads/' + branchControl.getText()

    let pullRequests = await getGitClient().getPullRequests(repoControl.getText(),
        <GitPullRequestSearchCriteria>{ targetRefName: targetBranch, status: PullRequestStatus.Completed }, VSS.getWebContext().project.id);

    pullRequests.forEach(req => {
        if (req.lastMergeCommit && commits.findIndex(x => x.commitId == req.lastMergeCommit.commitId) != -1) {
            reqResultArr.push({
                repo: repoControl.getText(), branch: branchTagLabel, date: req.creationDate,
                title: req.title, desc: req.description, id: req.pullRequestId
            });
            pullreqInGrid.push(req);
        };
    });
    reqGridControl.setDataSource(reqResultArr);
    reqResultArr = [];


    commits.forEach(commit => {
        resultArr.push({ repo: repoControl.getText(), branch: branchTagLabel, date: commit.author.date, commit: commit.comment, id: commit.commitId });
    });
    gridControl.setDataSource(resultArr);
    resultArr = [];
    waitControler.endWait();
    return pullreqInGrid;
};
// get commits between two tags.
async function getBatchCommits(firstTag, lastTag) {
    let alltags = await getTags();
    let filtedCommits: GitCommitRef[] = [];
    let commitsFilter;

    if (alltags.findIndex(x => x == firstTag) > alltags.findIndex(x => x == lastTag)) {
        return filtedCommits;
    }

    if (fromTagControl.getText() == alltags[0]) {
        commitsFilter = <GitQueryCommitsCriteria>{
            itemVersion: <GitVersionDescriptor>{ version: lastTag, versionType: GitVersionType.Tag },
            fromDate: startDateControl.getText(), toDate: nextDay(endDateControl.getText())
        };
        return await getGitClient().getCommits(repoControl.getText(), commitsFilter, VSS.getWebContext().project.id)
    }

    let firstTagIndex = alltags.findIndex(x => x == firstTag);

    if (firstTagIndex != 0) {
        firstTag = alltags[firstTagIndex - 1];
    }

    commitsFilter = <GitQueryCommitsCriteria>{
        itemVersion: <GitVersionDescriptor>{ version: firstTag, versionType: GitVersionType.Tag },
        fromDate: startDateControl.getText(), toDate: nextDay(endDateControl.getText())
    };

    let currentFromCommits = await getGitClient().getCommits(repoControl.getText(), commitsFilter, VSS.getWebContext().project.id);

    commitsFilter = <GitQueryCommitsCriteria>{
        itemVersion: <GitVersionDescriptor>{ version: lastTag, versionType: GitVersionType.Tag },
        fromDate: startDateControl.getText(), toDate: nextDay(endDateControl.getText())
    };
    let currentToCommits = await getGitClient().getCommits(repoControl.getText(), commitsFilter, VSS.getWebContext().project.id);


    currentToCommits.forEach(commit => {
        if (currentFromCommits.findIndex(x => x.commitId == commit.commitId) == -1) {
            filtedCommits.push(commit);
        }
    });

    return filtedCommits;
}

//formatting the date to fit data type in the combo control
function formatDate(date) {
    var day = date.getDate();
    var month = date.getMonth() + 1;
    var year = date.getFullYear();

    return month + '/' + day + '/' + year;
}

function nextDay(date) {
    var arr = date.split('/');
    arr[1] = (parseInt(arr[1]) + 1).toString();
    return arr[0] + '/' + arr[1] + '/' + arr[2];
}

function getPullrequestOfCommit(pullrequest: GitPullRequest, commit: GitCommitRef) {
    if (pullrequest.commits.findIndex((a) => a.commitId == commit.commitId) != -1) {
        return true
    }
    else
        return false;
}

async function getPullData(requestID: number) {
    var pullreq = await getGitClient().getPullRequest(repoControl.getText(), requestID, VSS.getWebContext().project.id, null, null, null, true);
    let elements = gridControl._canvas.children().contents().toArray();

    for (let i = 3; i < elements.length; i = i + 5) {
        $(elements[i].parentElement).css("background-color", "");
    }
    if (pullreq.commits) {
        pullreq.commits.forEach(commit => {
            for (let i = 3; i < elements.length; i = i + 5) {
                if (elements[i].innerHTML == commit.commitId) {
                    $(elements[i].parentElement).css("background-color", "DeepSkyBlue");
                }
            }
        });
    };
    return pullreq;
}

async function createTree() {
    if ($('.releaseTable').is(":visible")) {
        $('.reqTable').show();
        $('.commentTables').show();
        $('.releaseTable').hide();
        $('.releaseBtn').html('Show Tree');
        treeControl.dispose();
        releaseWaitControler.endWait();
        return;
    }
    else {
        treeControl = <TreeView.TreeView>BaseControl.createIn(TreeView.TreeView, $('.releaseTable'));
        $('.reqTable').hide();
        $('.commentTables').hide();
        $('.releaseTable').show();
        $('.releaseBtn').html('Hide Tree');
    }
    if (repoControl.getText() == '') {
        releaseWaitControler.endWait();
        return;
    };
    let source = await getRelaseData();
    // Converts the source to TreeNodes
    function convertToTreeNodes(items) {
        return $.map(items, function (item) {
            var node = new TreeView.TreeNode(item.name);
            node.icon = item.icon;
            node.expanded = item.expanded;
            if (item.children && item.children.length > 0) {
                node.addRange(convertToTreeNodes(item.children));
            }
            return node;
        });
    }

    // Generate TreeView options
    var treeviewOptions = {
        width: 400,
        height: "100%",
        nodes: convertToTreeNodes(source),
        clickToggles: true,
        clickSelects: true,
        useBowtieStyle: true
    };

    if (!treeControl.rootNode.hasChildren()) {
        treeControl.setEnhancementOptions(treeviewOptions);
        let nodes = convertToTreeNodes(source)
        treeControl.rootNode.addRange(nodes);
        treeControl.updateNode(treeControl.rootNode);
    }
    releaseWaitControler.endWait();
}
//crating an object with the required data for the treeview. 
async function getRelaseData() {
    let reqCommits: GitPullRequest[] = [];
    let WIs: WorkItem[] = [];
    let targetBranch = '';
    if (branchControl.getText() != '')
        targetBranch = 'refs/heads/' + branchControl.getText();
    let pullRequests = await getGitClient().getPullRequests(repoControl.getText(),
        <GitPullRequestSearchCriteria>{ targetRefName: targetBranch, status: PullRequestStatus.All }, VSS.getWebContext().project.id);
    let currentPullRequests = await getCommits();

    for (let i = 0; i < currentPullRequests.length; i++) {
        var pullreq = await getGitClient().getPullRequest(repoControl.getText(), currentPullRequests[i].pullRequestId, VSS.getWebContext().project.id, null, null, null, null, true);
        reqCommits.push(pullreq);
        if (pullreq.workItemRefs) {
            let p = [];
            for (let j = 0; j < pullreq.workItemRefs.length; j++) {
                p.push(getWIClient().getWorkItem(+pullreq.workItemRefs[j].id, null, null, wIContracts.WorkItemExpand.Relations).then((wi) => {
                    if (WIs.findIndex(x => x.id == wi.id) == -1) {
                        if (wi.fields['System.State'] == 'Done')
                            WIs.push(wi);
                    }
                }));
            }
            await Promise.all(p);
        }
    };
    let source = [];
    for (let i = 0; i < WIs.length; i++) {
        if (WIs[i].relations) {
            let children = [];
            for (let j = 0; j < WIs[i].relations.length; j++) {
                if (WIs[i].relations[j].attributes['name'] == 'Pull Request') {
                    let reqChildren = [];
                    let splittedArr = WIs[i].relations[j].url.toLocaleLowerCase().split('%2f');
                    let reqId = splittedArr[splittedArr.length - 1];
                    let currentReq = pullRequests.find(x => x.pullRequestId == +reqId);
                    if (currentReq) {
                        currentReq = await getGitClient().getPullRequest(repoControl.getText(), currentReq.pullRequestId, VSS.getWebContext().project.id, null, null, null, true);
                        if (currentReq.commits) {
                            currentReq.commits.forEach(commit => {
                                reqChildren.push({ name: commit.comment })
                            });
                        }
                        children.push({ name: currentReq.title + ' ' + reqId, children: reqChildren });
                    }
                }
            };
            source.push({ name: WIs[i].fields['System.Title'] + ' ' + WIs[i].id + ' ' + '(' + WIs[i].fields['System.WorkItemType'] + ')', children })
        }
    };
    return source;
};

async function createMarkDown() {
    if (repoControl.getText() == '') {
        releaseWaitControler.endWait();
        return;
    };
    if (toTagControl.getText() == '' || fromTagControl.getText() == '') {
        releaseWaitControler.endWait();
        alert("Please choose fill the 'From Tag' and 'To Tag' fields.")
        return;
    };

    let allTags = await getTags();
    let fromTag = fromTagControl.getText();
    let toTag = toTagControl.getText();
    let reqTags = await traverseTags(allTags, fromTag, toTag);
    let source = await getRelaseData();


    let firstFind = true;
    let htmlContent: string = '';
    htmlContent += ('<!DOCTYPE html>');
    htmlContent += ('<html>');
    htmlContent += ('<title>Relase Notes</title>');
    htmlContent += ('<body>');
    htmlContent += ('<div style="float: left">');
    htmlContent += ('<h1>' + repoControl.getText() + '</h1>');
    htmlContent += ('<ul>');
    reqTags.forEach(tag => {
        htmlContent += ('<h2>' + tag.key + '</h2>');
        htmlContent += ('<ul>');
        source.forEach(element => {
            tag['value'].forEach(req => {
                if (firstFind && element.children.findIndex(x => x.name.includes(req.pullRequestId))) {
                    htmlContent += ('<li>' + element.name + '</li>');
                    firstFind = false;
                };
            });
            firstFind = true;
        });
        htmlContent += ('</ul>');
    });
    htmlContent += ('</div>')

    let markContent: string = '';
    markContent += ('# ' + repoControl.getText() + "\n");
    reqTags.forEach(tag => {
        markContent += ('\n## ' + tag.key + "\n");
        source.forEach(element => {
            tag['value'].forEach(req => {
                //filtering only the relevent data.
                if (firstFind && element.children.findIndex(x => x.name.includes(req.pullRequestId))) {
                    markContent += ('* ' + element.name + "\n");
                    firstFind = false;
                };
            });
            firstFind = true;
        });
    });
    htmlContent += ('<div style="float: left;margin-left:10%">')
    htmlContent += ('<h1>MarkDown</h1>');
    htmlContent += ('<textarea style="width:600px; height:900px;overflow:auto;">');
    htmlContent += markContent;
    htmlContent += ('</textarea>');
    htmlContent += ('</div>')
    htmlContent += ('</body>');
    htmlContent += ('</html>');

    let w = window.open("", "popupWindow", "width=" + $(window).width() + ", height=" + $(window).height() + ", scrollbars=yes");
    var $w = $(w.document.body);
    $w.html(htmlContent);
    releaseWaitControler.endWait();
};

async function traverseTags(tags: string[], fromTag, toTag) {
    let firstTagIndex = tags.findIndex(x => x == fromTag);
    let lastTagIndex = tags.findIndex(x => x == toTag);
    let tagDict = [];
    for (let i = firstTagIndex; i <= lastTagIndex; i++) {
        let commits = await getBatchCommits(tags[i], tags[i])
        tagDict.push({ key: tags[i], value: await getPullRequestsFromComments(commits) });
    };

    return tagDict;
};


async function getPullRequestsFromComments(commits: GitCommitRef[]) {
    let targetBranch;
    let reqs = [];

    if (branchControl.getText() != '')
        targetBranch = 'refs/heads/' + branchControl.getText()
    let pullRequests = await getGitClient().getPullRequests(repoControl.getText(),
        <GitPullRequestSearchCriteria>{ targetRefName: targetBranch, status: PullRequestStatus.Completed }, VSS.getWebContext().project.id);
    pullRequests.forEach(req => {
        if (req.lastMergeCommit && commits.findIndex(x => x.commitId == req.lastMergeCommit.commitId) != -1) {
            reqs.push(req);
        };
    });
    return reqs;
}
