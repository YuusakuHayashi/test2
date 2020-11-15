2020/11/14 問題点 : ViewModel.ViewsとViewExplorerViewModel.Viewsの同期
    ViewModel.Viewsへ直接オブジェクト(ObservableCollection(Of ViewItemModel))がセットされたときに、
    DelegateEventListener.Instance.RaiseViewChanged()を発生させて、購読しているViewExplorerViewModelの
    Viewsプロパティと同期をとっている。
    だが、ViewModel.CreateViewsでViewModel.Viewsを作成するとき、これは直接セットしていない
    (ViewModel.Views.Addしている)にも関わらず、ViewExplorerViewModel.Viewsと同期している
    これはオブジェクトのため、vm.Viewsとvevm.Viewsは参照関係にあるからだと思う？が、
    DelegateEventListener.Instance.RaiseViewChanged()を無効にしてしまうと、うまく同期しなくなる
    原因は不明だが、Sub()でViewModel.Viewを直接更新するのではなく、Function()にして、
    ViewModel.Views = _CreateViews(fvm, Me.Views)とした。

