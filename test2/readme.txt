2020/11/14 ������ : ViewModel.Views��ViewExplorerViewModel.Views��Ʊ��
    ViewModel.Views��ľ�ܥ��֥�������(ObservableCollection(Of ViewItemModel))�����åȤ��줿�Ȥ��ˡ�
    DelegateEventListener.Instance.RaiseViewChanged()��ȯ�������ơ����ɤ��Ƥ���ViewExplorerViewModel��
    Views�ץ�ѥƥ���Ʊ����ȤäƤ��롣
    ������ViewModel.CreateViews��ViewModel.Views���������Ȥ��������ľ�ܥ��åȤ��Ƥ��ʤ�
    (ViewModel.Views.Add���Ƥ���)�ˤ�ؤ�餺��ViewExplorerViewModel.Views��Ʊ�����Ƥ���
    ����ϥ��֥������ȤΤ��ᡢvm.Views��vevm.Views�ϻ��ȴط��ˤ��뤫����Ȼפ�������
    DelegateEventListener.Instance.RaiseViewChanged()��̵���ˤ��Ƥ��ޤ��ȡ����ޤ�Ʊ�����ʤ��ʤ�
    ����������������Sub()��ViewModel.View��ľ�ܹ�������ΤǤϤʤ���Function()�ˤ��ơ�
    ViewModel.Views = _CreateViews(fvm, Me.Views)�Ȥ�����

