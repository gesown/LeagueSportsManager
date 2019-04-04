var ViewModel = function (first, last) {
    var self = this;
    self.firstName = ko.observable(first);
    self.lastName = ko.observable(last);
    self.fullName = ko.computed(function () {
        return self.firstName() + " " + self.lastName();
    }, self);
};
var RegisterViewModel = function(registered) {
    var self = this;
    self.notRegistered = ko.observable(registered);
    return self;
};
function HomeViewModel(app, dataModel) {
    var self = this;

    self.myHometown = ko.observable("");
    self.ViewModel = ko.observable(ViewModel("marty", "grogan"));
    self.RegisterViewModel = ko.observable(RegisterViewModel(true));
    Sammy(function () {
        this.get('#home', function () {
            // Make a call to the protected Web API by passing in a Bearer Authorization Header
            $.ajax({
                method: 'get',
                url: app.dataModel.userInfoUrl,
                contentType: "application/json; charset=utf-8",
                headers: {
                    'Authorization': 'Bearer ' + app.dataModel.getAccessToken()
                },
                success: function (data) {
                    self.myHometown('Your Hometown is : ' + data.hometown);
                }
            });
        });
        this.get('/', function () { this.app.runRoute('get', '#home') });
    });

    return self;
}

app.addViewModel({
    name: "Home",
    bindingMemberName: "home",
    factory: HomeViewModel
});
app.addViewModel({
    name: "ViewModel",
    bindingMemberName: "registerMain",
    factory: HomeViewModel

});
app.addViewModel({
    name: "RegisterViewModel",
    bindingMemberName: "registerDiv",
    factory: HomeViewModel
});
