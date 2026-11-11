// FilePropsTool — 向导式: 1 选择文件 → 2 文件系统时间 → 3 Office 属性(不支持则跳过) → 4 保存
// 文件系统时间: SetFileTime (创建/修改/访问)
// Office 属性: docProps/core.xml 的 dcterms:created / dcterms:modified / cp:lastPrinted
//   XML 存 UTC，界面显示本地时间；写入时本地→UTC 转换
// build: build.bat (VS2019 x64, MFC 动态链接, miniz 处理 zip)
#include <afxwin.h>
#include <afxdtctl.h> // CDateTimeCtrl
#include <afxcmn.h>  // CToolTipCtrl
#include <afxdlgs.h>  // CFileDialog
#include "miniz.h"
#include <string>
#include <vector>
#include <cstdio>
#pragma comment(lib, "comctl32.lib")
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#define IDC_PATH    1001
#define IDC_PICK    1002
#define IDC_CHKFS   1010
#define IDC_FS1     1011
#define IDC_FS2     1012
#define IDC_FS3     1013
#define IDC_CHKDOC  1020
#define IDC_DC1     1021
#define IDC_DC2     1022
#define IDC_DC3     1023
#define IDC_STATUS  1031
#define IDC_STATUS  1031
#define IDC_TX1     1050   // 作者..经理 六个文本属性
#define IDC_TX2     1051
#define IDC_TX3     1052
#define IDC_TX4     1053
#define IDC_TX5     1054
#define IDC_TX6     1055
#define IDC_BACK    1040
#define IDC_NEXT    1041
#define IDC_RECENT  1044
#define IDC_RX     1100   // 最近记录行 x 按钮 (1100-1104)

// 带滚动通知的 ListBox: 滚动时让父窗口重对位 x 按钮
class CXList : public CListBox
{
    DECLARE_MESSAGE_MAP()
public:
    afx_msg void VScroll(UINT, UINT);
};
BEGIN_MESSAGE_MAP(CXList, CListBox)
    ON_WM_VSCROLL_REFLECT()
END_MESSAGE_MAP()

void CXList::VScroll(UINT, UINT)
{
    GetParent()->PostMessageW(WM_APP + 1);
}

// ---------- ISO8601(UTC) <-> SYSTEMTIME ----------
static bool IsoToUtc(const std::string& s, SYSTEMTIME& st)
{
    int y, mo, d, h, mi, se;
    if (sscanf_s(s.c_str(), "%4d-%2d-%2dT%2d:%2d:%2d", &y, &mo, &d, &h, &mi, &se) != 6)
        return false;
    st = { (WORD)y, (WORD)mo, 0, (WORD)d, (WORD)h, (WORD)mi, (WORD)se, 0 };
    return true;
}
static std::string UtcToIso(const SYSTEMTIME& st)
{
    char buf[40];
    sprintf_s(buf, "%04u-%02u-%02uT%02u:%02u:%02uZ", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    return buf;
}
static bool UtcToLocal(const SYSTEMTIME& utc, SYSTEMTIME& loc)
{
    SYSTEMTIME u = utc;
    return !!SystemTimeToTzSpecificLocalTime(NULL, &u, &loc);
}
static bool LocalToUtc(const SYSTEMTIME& loc, SYSTEMTIME& utc)
{
    SYSTEMTIME l = loc;
    return !!TzSpecificLocalTimeToSystemTime(NULL, &l, &utc);
}

// ---------- core.xml 标签内文本读写 ----------
// <tag ...>inner</tag> —— 只换 inner，保留属性
// <tag ...>inner</tag> —— 只换 inner，保留属性; 兼容 <tag/> 自闭合
static bool GetInner(const std::string& xml, const char* tag, std::string& out)
{
    std::string open = "<"; open += tag;
    size_t p = xml.find(open);
    if (p == std::string::npos) return false;
    size_t gt = xml.find('>', p);
    if (gt == std::string::npos) return false;
    if (xml[gt - 1] == '/') { out.clear(); return true; }   // <tag/>
    size_t e = xml.find('<', gt + 1);
    if (e == std::string::npos) return false;
    out = xml.substr(gt + 1, e - gt - 1);
    return true;
}
static bool ReplaceInner(std::string& xml, const char* tag, const std::string& val)
{
    std::string open = "<"; open += tag;
    size_t p = xml.find(open);
    if (p == std::string::npos) return false;
    size_t gt = xml.find('>', p);
    if (gt == std::string::npos) return false;
    if (xml[gt - 1] == '/')
    {   // <tag/> -> <tag>val</tag>
        std::string full = open + ">" + val + "</" + tag + ">";
        xml.replace(p, gt + 1 - p, full);
        return true;
    }
    size_t e = xml.find('<', gt + 1);
    if (e == std::string::npos) return false;
    xml.replace(gt + 1, e - gt - 1, val);
    return true;
}
// 无该标签时在 endTag 前插入 <tag>val</tag>
static bool SetTag(std::string& xml, const char* tag, const std::string& val, const char* endTag)
{
    if (ReplaceInner(xml, tag, val)) return true;
    if (val.empty()) return true;                            // 原本没有且值为空: 不新增
    size_t p = xml.find(endTag);
    if (p == std::string::npos) return false;
    std::string ins = "<" + std::string(tag) + ">" + val + "</" + tag + ">";
    xml.insert(p, ins);
    return true;
}
static CStringW FromUtf8(const std::string& in);
static std::string ToUtf8(const CStringW& in)   // XML 是 UTF-8, 不能走 ANSI 码页
{
    int n = WideCharToMultiByte(CP_UTF8, 0, in, -1, NULL, 0, NULL, NULL);
    std::string s(n > 1 ? n - 1 : 0, '\0');
    if (n > 1) WideCharToMultiByte(CP_UTF8, 0, in, -1, &s[0], n, NULL, NULL);
    return s;
}
static std::string XmlEscape(const CStringW& in)
{
    std::string s = ToUtf8(in), o;
    for (char c : s)
    {
        switch (c)
        {
        case '&': o += "&amp;"; break;
        case '<': o += "&lt;"; break;
        case '>': o += "&gt;"; break;
        case '"': o += "&quot;"; break;
        case '\'': o += "&apos;"; break;
        default: o += c;
        }
    }
    return o;
}
static CStringW XmlUnescape(const std::string& in)
{
    std::string s = in, o;
    for (size_t i = 0; i < s.size(); i++)
    {
        if (s[i] == '&')
        {
            size_t sc = s.find(';', i);
            if (sc != std::string::npos && sc - i <= 6)
            {
                std::string ent = s.substr(i + 1, sc - i - 1);
                if (ent == "amp") { o += '&'; i = sc; continue; }
                if (ent == "lt") { o += '<'; i = sc; continue; }
                if (ent == "gt") { o += '>'; i = sc; continue; }
                if (ent == "quot") { o += '"'; i = sc; continue; }
                if (ent == "apos") { o += '\''; i = sc; continue; }
            }
        }
        o += s[i];
    }
    return FromUtf8(o);
}
static CStringW FromUtf8(const std::string& in)
{
    int n = MultiByteToWideChar(CP_UTF8, 0, in.c_str(), (int)in.size(), NULL, 0);
    CStringW w;
    if (n > 0)
    {
        wchar_t* p = w.GetBuffer(n);
        MultiByteToWideChar(CP_UTF8, 0, in.c_str(), (int)in.size(), p, n);
        w.ReleaseBuffer(n);
    }
    return w;
}

// ---------- 最近记录: %LOCALAPPDATA%\FilePropsTool\recent.txt (UTF-8, 每行一路径, 新的在前) ----------
static CStringW RecentPath()
{
    wchar_t buf[MAX_PATH] = { 0 };
    ::GetEnvironmentVariableW(L"LOCALAPPDATA", buf, MAX_PATH);
    CStringW dir(buf);
    dir += L"\\FilePropsTool";
    ::CreateDirectoryW(dir, NULL);
    return dir + L"\\recent.txt";
}
static std::vector<CStringW> LoadRecent()
{
    std::vector<CStringW> v;
    FILE* f = _wfopen(RecentPath(), L"rb");
    if (!f) return v;
    std::string s;
    char buf[4096]; size_t n;
    while ((n = fread(buf, 1, sizeof(buf), f)) > 0) s.append(buf, n);
    fclose(f);
    CStringW all = FromUtf8(s);
    int cur = 0;
    for (int i = 0; i <= all.GetLength(); i++)
    {
        if (i == all.GetLength() || all[i] == L'\n')
        {
            CStringW line = all.Mid(cur, i - cur);
            line.Trim(L"\r \"");
            if (!line.IsEmpty()) v.push_back(line);
            cur = i + 1;
        }
    }
    if (v.size() > 5) v.resize(5);
    return v;
}
static void StoreRecent(const std::vector<CStringW>& v)
{
    FILE* f = _wfopen(RecentPath(), L"wb");
    if (!f) return;
    for (const auto& p : v)
    {
        std::string line = ToUtf8(p) + "\n";
        fwrite(line.data(), 1, line.size(), f);
    }
    fclose(f);
}
static void AddRecent(const CStringW& path)
{
    std::vector<CStringW> v = LoadRecent();
    for (auto it = v.begin(); it != v.end(); )
        if (it->CompareNoCase(path) == 0) it = v.erase(it); else ++it;
    v.insert(v.begin(), path);
    if (v.size() > 5) v.resize(5);
    StoreRecent(v);
}
class CMainFrame : public CFrameWnd
{
    DECLARE_MESSAGE_MAP()
public:
    afx_msg int OnCreate(LPCREATESTRUCT);
    virtual BOOL PreTranslateMessage(MSG* pMsg);
    void UpdatePathTip();   // 路径显示不全时启用 tooltip
    afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
    afx_msg void OnDropFiles(HDROP);
    afx_msg void OnPick();
    afx_msg void OnRecentSel();
    afx_msg void OnRecentX(UINT nID);
    afx_msg LRESULT OnReposX(WPARAM, LPARAM);
    afx_msg void OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDIS);
    afx_msg void OnBack();
    afx_msg void OnNext();
private:
    void LoadFileData(const CStringW& path);
    void RefreshRecent();
    bool ApplyChanges();
    bool WriteCoreXml(const CStringW& path, const std::string* core, const std::string* app);
    void EnableDocGroup(bool on);
    void PositionXButtons();
    void SetPage(int n);
    void Say(const wchar_t* fmt, ...);

    CFont     m_font;
    CBrush    m_brWhite;
    CFont     m_fontDrop;
    CEdit     m_editPath;
    CXList    m_listRecent;
    CToolTipCtrl m_tip;      // 路径过长时的完整路径提示
    CStatic   m_stRecent;
    CButton   m_btnPick, m_btnBack, m_btnNext;
    CButton   m_btnX[5];   // 每条记录前的 x 删除按钮
    CDateTimeCtrl m_fs[3];   // 创建/修改/访问 (文件系统)
    CDateTimeCtrl m_dc[3];   // 创建/修改/打印 (Office 属性)
    CStatic   m_stStatus, m_stWelcome, m_stDocInfo;
    CEdit     m_tx[6];     // 作者..经理
    std::vector<CWnd*> m_pgWelcome, m_pgFs, m_pgDoc;
    CStringW  m_path;
    bool      m_hasDoc = false;
    std::string m_core;      // 原始 core.xml (UTF-8)
    std::string m_appCore;   // docProps/app.xml (公司/经理)
    bool      m_hasApp = false;
    int       m_page = -1;
    bool      m_saving = false;   // 保存防重入
};

BEGIN_MESSAGE_MAP(CMainFrame, CFrameWnd)
    ON_WM_CREATE()
    ON_WM_CTLCOLOR()
    ON_WM_DROPFILES()
    ON_BN_CLICKED(IDC_PICK, &CMainFrame::OnPick)
    ON_LBN_SELCHANGE(IDC_RECENT, &CMainFrame::OnRecentSel)
    ON_CONTROL_RANGE(BN_CLICKED, IDC_RX, IDC_RX + 4, &CMainFrame::OnRecentX)
    ON_MESSAGE(WM_APP + 1, &CMainFrame::OnReposX)
    ON_WM_DRAWITEM()
    ON_BN_CLICKED(IDC_BACK, &CMainFrame::OnBack)
    ON_BN_CLICKED(IDC_NEXT, &CMainFrame::OnNext)
END_MESSAGE_MAP()

int CMainFrame::OnCreate(LPCREATESTRUCT cs)
{
    if (CFrameWnd::OnCreate(cs) == -1) return -1;

    NONCLIENTMETRICS ncm = { sizeof(ncm) };
    SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
    m_font.CreateFontIndirect(&ncm.lfMessageFont);
    m_brWhite.CreateSolidBrush(RGB(255, 255, 255));

    auto lbl = [&](const wchar_t* t, int x, int y, int w, std::vector<CWnd*>& page) {
        CStatic* s = new CStatic();
        s->Create(t, WS_CHILD, CRect(x, y, x + w, y + 20), this);
        s->SetFont(&m_font); page.push_back(s); return s; };
    auto dtp = [&](int id, int x, int y, std::vector<CWnd*>& page) {
        int i = (id >= IDC_DC1) ? id - IDC_DC1 : id - IDC_FS1;
        CDateTimeCtrl* p = &((id >= IDC_DC1) ? m_dc[i] : m_fs[i]);
        p->Create(WS_CHILD | WS_TABSTOP | DTS_TIMEFORMAT,
                  CRect(x, y - 4, x + 210, y + 18), this, id);   // 与文本输入框同宽同高
        p->SetFormat(_T("yyyy-MM-dd HH:mm:ss"));
        p->SetFont(&m_font); page.push_back(p); return p; };
    auto grp = [&](const wchar_t* t, int y, int h, std::vector<CWnd*>& page) {
        CButton* g = new CButton();
        g->Create(t, WS_CHILD | BS_GROUPBOX, CRect(10, y, 655, y + h), this, 0x9000);
        g->SetFont(&m_font); page.push_back(g); return g; };

    // 顶栏(所有页可见): 文件路径
    CRect rcClient;
    GetClientRect(&rcClient);
    m_editPath.Create(WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL | ES_READONLY,
                      CRect(12, 12, rcClient.Width() - 12, 36), this, IDC_PATH);
    m_editPath.SetFont(&m_font);
    m_editPath.ModifyStyleEx(0, WS_EX_CLIENTEDGE);   // 输入框边框
    m_tip.Create(this, TTS_NOPREFIX);
    m_tip.AddTool(&m_editPath, _T(""));
    m_tip.Activate(FALSE);

    // 第1页 拖放区(上半) + 最近文件(下半)
    m_btnPick.Create(_T("拖入文件到这里，或点击选择文件"),
                     WS_CHILD | WS_TABSTOP | BS_OWNERDRAW,
                     CRect(10, 56, rcClient.Width() - 10, 240), this, IDC_PICK);
    m_pgWelcome.push_back(&m_btnPick);
    m_stRecent.Create(_T("最近使用："), WS_CHILD, CRect(10, 248, 186, 268), this);
    m_stRecent.SetFont(&m_font);
    m_pgWelcome.push_back(&m_stRecent);
    m_listRecent.Create(WS_CHILD | WS_VSCROLL | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT,
                        CRect(34, 268, rcClient.Width() - 24, 364), this, IDC_RECENT);
    m_listRecent.SetFont(&m_font);
    m_pgWelcome.push_back(&m_listRecent);
    for (int i = 0; i < 5; i++)
    {
        m_btnX[i].Create(_T("✕"), WS_CHILD | BS_OWNERDRAW, CRect(0, 0, 18, 16), this, IDC_RX + i);
    }
    LOGFONT lf;
    m_font.GetLogFont(&lf);
    lf.lfHeight = -17;
    lf.lfWeight = FW_SEMIBOLD;
    m_fontDrop.CreateFontIndirect(&lf);
    m_btnPick.SetFont(&m_fontDrop);

    // 第2页 文件系统时间
    lbl(_T("创建时间"), 40, 160, 70, m_pgFs);  dtp(IDC_FS1, 130, 160, m_pgFs);
    lbl(_T("修改时间"), 40, 184, 70, m_pgFs);  dtp(IDC_FS2, 130, 184, m_pgFs);
    lbl(_T("访问时间"), 40, 208, 70, m_pgFs);  dtp(IDC_FS3, 130, 208, m_pgFs);

    // 第3页 Office 属性 (顺序同 Windows 详细信息: 作者/上次保存者/修订号/版本/公司/经理/日期)
    auto txt = [&](int id, const wchar_t* t, int y) {
        lbl(t, 40, y, 76, m_pgDoc);
        CEdit* e = &m_tx[id - IDC_TX1];
        e->Create(WS_CHILD | WS_TABSTOP | ES_AUTOHSCROLL, CRect(130, y - 4, 340, y + 18), this, id);
        e->ModifyStyleEx(0, WS_EX_CLIENTEDGE);        // 输入框边框
        e->SetFont(&m_font); m_pgDoc.push_back(e); };
    txt(IDC_TX1, _T("作者"),      90);   // dc:creator
    txt(IDC_TX2, _T("上次保存者"), 114); // cp:lastModifiedBy
    txt(IDC_TX3, _T("修订号"),    138);  // cp:revision
    txt(IDC_TX4, _T("版本号"),    162);  // cp:version
    txt(IDC_TX5, _T("公司"),      186);  // app.xml Company
    txt(IDC_TX6, _T("经理"),      210);  // app.xml Manager
    lbl(_T("创建时间"), 40, 234, 70, m_pgDoc);  dtp(IDC_DC1, 130, 234, m_pgDoc);
    lbl(_T("修改时间"), 40, 258, 70, m_pgDoc);  dtp(IDC_DC2, 130, 258, m_pgDoc);
    lbl(_T("打印时间"), 40, 282, 70, m_pgDoc);  dtp(IDC_DC3, 130, 282, m_pgDoc);


    // 底栏(所有页可见): 导航 + 状态
    m_btnBack.Create(_T("< 上一步"), WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                     CRect(440, 374, 528, 402), this, IDC_BACK);
    m_btnBack.SetFont(&m_font);
    m_btnNext.Create(_T("下一步 >"), WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
                     CRect(538, 374, 645, 402), this, IDC_NEXT);
    m_btnNext.SetFont(&m_font);
    m_stStatus.Create(_T(""), WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS,
                      CRect(12, 384, 420, 404), this, IDC_STATUS);
    m_stStatus.SetFont(&m_font);

    DragAcceptFiles(TRUE);
    EnableDocGroup(false);
    SetPage(0);
    RefreshRecent();

    CWinApp* app = (CWinApp*)AfxGetApp();
    if (app->m_lpCmdLine && app->m_lpCmdLine[0])
    {
        CStringW p(app->m_lpCmdLine);
        p.Trim(_T("\" "));
        if (!p.IsEmpty()) { LoadFileData(p); SetPage(1); }
    }
    return 0;
}

void CMainFrame::SetPage(int n)
{
    m_page = n;
    auto show = [](const std::vector<CWnd*>& v, BOOL on) {
        for (CWnd* w : v) w->ShowWindow(on); };
    show(m_pgWelcome, n == 0);
    show(m_pgFs, n == 1);
    if (n != 0) for (auto& b : m_btnX) b.ShowWindow(SW_HIDE);   // x 按钮不在页向量, 切屏手动隐藏
    else PositionXButtons();                                    // 回首页: 重新对位显示
    show(m_pgDoc, n == 2);
    m_btnBack.EnableWindow(n > 0);
    bool last = (n == 2 && m_hasDoc) || (n == 1 && !m_hasDoc);
    m_btnNext.SetWindowText(last ? _T("保存") : _T("下一步 >"));
    m_btnNext.EnableWindow(n > 0 ? TRUE : !m_path.IsEmpty());
}


void CMainFrame::EnableDocGroup(bool on)
{
    for (auto& d : m_dc) d.EnableWindow(on);
    for (auto& e : m_tx) e.EnableWindow(on);
}

HBRUSH CMainFrame::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
    // 统一白底; 禁用浅灰字提示不可编辑, 启用黑字
    // 投放区按钮: 不擦背景(自绘整块), 消除闪烁
    if (pWnd->GetDlgCtrlID() == IDC_PICK)
        return (HBRUSH)::GetStockObject(NULL_BRUSH);
    pDC->SetTextColor(pWnd->IsWindowEnabled() ? RGB(0, 0, 0) : RGB(176, 176, 176));
    pDC->SetBkColor(RGB(255, 255, 255));
    return (HBRUSH)m_brWhite.GetSafeHandle();
}

void CMainFrame::OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT di)
{
    if (nIDCtl >= IDC_RX && nIDCtl < IDC_RX + 5)
    {   // 行内删除按钮: 同底色 + 红色 ✕ (居中方形, 不挤压)
        CDC* pDC = CDC::FromHandle(di->hDC);
        CRect rc(di->rcItem);
        pDC->FillSolidRect(rc, RGB(255, 255, 255));
        int cx = (rc.left + rc.right) / 2, cy = (rc.top + rc.bottom) / 2;
        int r = (rc.Width() < rc.Height() ? rc.Width() : rc.Height()) / 2 - 2;
        CPen pen(PS_SOLID, 1, RGB(200, 30, 30));
        CPen* op = pDC->SelectObject(&pen);
        pDC->MoveTo(cx - r, cy - r); pDC->LineTo(cx + r, cy + r);
        pDC->MoveTo(cx + r, cy - r); pDC->LineTo(cx - r, cy + r);
        pDC->SetPixelV(cx + r, cy + r, RGB(200, 30, 30));  // 补上 LineTo 不画的终点, 保证上下对称
        pDC->SetPixelV(cx - r, cy + r, RGB(200, 30, 30));
        pDC->SelectObject(op);
        return;
    }
    if (nIDCtl != IDC_PICK) { CFrameWnd::OnDrawItem(nIDCtl, di); return; }
    CDC* pDC = CDC::FromHandle(di->hDC);
    CRect rc(di->rcItem);
    COLORREF bg = (di->itemState & ODS_SELECTED) ? RGB(238, 242, 248) : RGB(250, 251, 253);
    pDC->FillSolidRect(rc, bg);
    CPen pen(PS_DASH, 1, RGB(170, 178, 190));
    CPen* op = pDC->SelectObject(&pen);
    pDC->SetBkMode(TRANSPARENT);
    pDC->Rectangle(rc);
    pDC->SelectObject(op);
    CFont* of = pDC->SelectObject(&m_fontDrop);
    pDC->SetTextColor(RGB(70, 80, 95));
    pDC->DrawText(_T("拖入文件到这里，或点击选择文件"), rc,
                  DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    pDC->SelectObject(of);
    if (di->itemState & ODS_FOCUS)
    {   // 焦点: 细内框, 不抢眼
        CPen fp(PS_SOLID, 1, RGB(150, 160, 175));
        op = pDC->SelectObject(&fp);
        pDC->Rectangle(rc);
        pDC->SelectObject(op);
    }
}

void CMainFrame::Say(const wchar_t* fmt, ...)
{
    wchar_t buf[512];
    va_list ap; va_start(ap, fmt);
    _vsnwprintf_s(buf, _TRUNCATE, fmt, ap);
    va_end(ap);
    m_stStatus.SetWindowText(buf);
}

void CMainFrame::OnDropFiles(HDROP h)
{
    wchar_t path[MAX_PATH];
    if (DragQueryFileW(h, 0, path, MAX_PATH))
    {
        LoadFileData(path);
        SetPage(1);
    }
    DragFinish(h);
}

void CMainFrame::OnPick()
{
    CFileDialog dlg(TRUE, NULL, NULL, OFN_FILEMUSTEXIST | OFN_HIDEREADONLY,
                    _T("所有文件 (*.*)|*.*||"), this);
    INT_PTR r = dlg.DoModal();
    m_btnPick.Invalidate(FALSE);        // 模态后强制重绘 owner-draw (comctl6 不主动发 WM_DRAWITEM)
    if (r == IDOK)
    {
        LoadFileData(dlg.GetPathName());
        SetPage(1);
    }
}

void CMainFrame::LoadFileData(const CStringW& path)
{
    m_path = path;
    m_core.clear(); m_hasDoc = false;
    m_appCore.clear(); m_hasApp = false;
    m_editPath.SetWindowText(path);
    UpdatePathTip();

    HANDLE h = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                           NULL, OPEN_EXISTING, 0, NULL);
    if (h == INVALID_HANDLE_VALUE) { Say(_T("无法打开文件: %lu"), GetLastError()); return; }
    FILETIME c, a, w;
    GetFileTime(h, &c, &a, &w);
    CloseHandle(h);

    auto setFs = [](CDateTimeCtrl& dtp, const FILETIME& ft) {
        SYSTEMTIME utc, loc;
        if (FileTimeToSystemTime(&ft, &utc) && UtcToLocal(utc, loc)) dtp.SetTime(&loc); };
    setFs(m_fs[0], c); setFs(m_fs[1], w); setFs(m_fs[2], a);

    // Office OpenXML: 尝试读 docProps/core.xml
    if (FILE* fz = _wfopen(path, L"rb"))
    {
        mz_zip_archive z = { 0 };
        if (mz_zip_reader_init_cfile(&z, fz, 0, 0))
        {
            int i = mz_zip_reader_locate_file(&z, "docProps/core.xml", NULL, 0);
            if (i >= 0)
            {
                mz_zip_archive_file_stat st;
                if (mz_zip_reader_file_stat(&z, i, &st) && st.m_uncomp_size && st.m_uncomp_size < 4u * 1024 * 1024)
                {
                    std::string buf((size_t)st.m_uncomp_size, '\0');
                    if (mz_zip_reader_extract_to_mem(&z, i, &buf[0], (size_t)st.m_uncomp_size, 0))
                    {
                        m_core.swap(buf);
                        m_hasDoc = true;
                    }
                }
            }
            int j = mz_zip_reader_locate_file(&z, "docProps/app.xml", NULL, 0);
            if (j >= 0)
            {
                mz_zip_archive_file_stat st;
                if (mz_zip_reader_file_stat(&z, j, &st) && st.m_uncomp_size && st.m_uncomp_size < 4u * 1024 * 1024)
                {
                    std::string buf((size_t)st.m_uncomp_size, '\0');
                    if (mz_zip_reader_extract_to_mem(&z, j, &buf[0], (size_t)st.m_uncomp_size, 0))
                    {
                        m_appCore.swap(buf);
                        m_hasApp = true;
                    }
                }
            }
            mz_zip_reader_end(&z);
        }
        fclose(fz);
    }

    if (m_hasDoc)
    {
        auto setDoc = [&](CDateTimeCtrl& dtp, const char* tag) {
            std::string v;
            SYSTEMTIME utc, loc;
            if (GetInner(m_core, tag, v) && IsoToUtc(v, utc) && UtcToLocal(utc, loc))
                dtp.SetTime(&loc); };
        setDoc(m_dc[0], "dcterms:created");
        setDoc(m_dc[1], "dcterms:modified");
        setDoc(m_dc[2], "cp:lastPrinted"); // 无该属性时保持当前值
        // 文本属性: core.xml 4 个 + app.xml 2 个
        auto setTx = [&](CEdit& e, const char* tag) {
            std::string v;
            e.SetWindowText(GetInner(m_core, tag, v) ? XmlUnescape(v) : L""); };
        setTx(m_tx[0], "dc:creator");
        setTx(m_tx[1], "cp:lastModifiedBy");
        setTx(m_tx[2], "cp:revision");
        setTx(m_tx[3], "cp:version");
        auto setTxApp = [&](CEdit& e, const char* tag) {
            std::string v;
            e.SetWindowText((m_hasApp && GetInner(m_appCore, tag, v)) ? XmlUnescape(v) : L""); };
        setTxApp(m_tx[4], "Company");
        setTxApp(m_tx[5], "Manager");
        EnableDocGroup(true);
    }
    else
    {
        EnableDocGroup(false);
        Say(_T("已加载"));
    }
}

void CMainFrame::OnRecentSel()
{
    int i = m_listRecent.GetCurSel();
    if (i < 0) return;
    CStringW p;
    m_listRecent.GetText(i, p);
    if (::GetFileAttributesW(p) == INVALID_FILE_ATTRIBUTES)
    {   // 文件已不存在: 从最近记录移除
        std::vector<CStringW> v = LoadRecent();
        for (auto it = v.begin(); it != v.end(); )
            if (it->CompareNoCase(p) == 0) it = v.erase(it); else ++it;
        StoreRecent(v);
        RefreshRecent();
        Say(_T("文件不存在，已从最近记录移除"));
        return;
    }
    LoadFileData(p);
    AddRecent(p);        // 点过的记录提到最前 (时间=点击时刻)
    RefreshRecent();
    SetPage(1);
}

void CMainFrame::OnRecentX(UINT nID)
{
    int i = nID - IDC_RX;
    std::vector<CStringW> v = LoadRecent();
    if (i < 0 || i >= (int)v.size()) return;
    v.erase(v.begin() + i);
    StoreRecent(v);
    RefreshRecent();
    Say(_T("已删除 1 条记录"));
}

void CMainFrame::UpdatePathTip()
{
    bool need = false;
    if (m_tip.m_hWnd && !m_path.IsEmpty())
    {
        CDC* dc = GetDC();
        CFont* of = dc->SelectObject(m_editPath.GetFont());
        CSize sz = dc->GetTextExtent(m_path);
        dc->SelectObject(of);
        ReleaseDC(dc);
        CRect rc;
        m_editPath.GetClientRect(&rc);
        need = sz.cx > rc.Width() - 10;   // 减去编辑框左右内边距
    }
    if (need) m_tip.UpdateTipText(m_path, &m_editPath);
    m_tip.Activate(need);
}

BOOL CMainFrame::PreTranslateMessage(MSG* pMsg)
{
    if (m_tip.m_hWnd && pMsg->message >= WM_MOUSEFIRST && pMsg->message <= WM_MOUSELAST)
        m_tip.RelayEvent(pMsg);
    return CFrameWnd::PreTranslateMessage(pMsg);
}

void CMainFrame::RefreshRecent()
{
    m_listRecent.ResetContent();
    for (const auto& p : LoadRecent())
        m_listRecent.AddString(p);
    BOOL show = m_listRecent.GetCount() > 0;
    m_stRecent.ShowWindow(show);
    m_listRecent.ShowWindow(show);
    for (auto& b : m_btnX) b.ShowWindow(show);
    PositionXButtons();
}

void CMainFrame::PositionXButtons()
{
    if (!m_listRecent.GetSafeHwnd()) return;
    int top = m_listRecent.GetTopIndex();
    int ih = m_listRecent.GetItemHeight(0);
    if (ih <= 0) ih = 18;
    CRect rcList;
    m_listRecent.GetWindowRect(&rcList);
    ScreenToClient(&rcList);
    int count = m_listRecent.GetCount();
    int visible = (rcList.Height() - 4) / ih;      // 一屏能放的行数
    for (int row = 0; row < 5; row++)
    {
        bool vis = row < count && row >= top && (row - top) < visible;
        if (vis)
        {
            int y = rcList.top + 2 + (row - top) * ih;
            m_btnX[row].MoveWindow(rcList.left - 24, y, 18, ih - 2);
            m_btnX[row].ShowWindow(SW_SHOW);
        }
        else
            m_btnX[row].ShowWindow(SW_HIDE);
    }
}

LRESULT CMainFrame::OnReposX(WPARAM, LPARAM)
{
    PositionXButtons();
    return 0;
}
bool CMainFrame::WriteCoreXml(const CStringW& path, const std::string* core, const std::string* app)
{
    CStringW tmp = path + L".tmp~";
    FILE* fi = _wfopen(path, L"rb");
    FILE* fo = _wfopen(tmp, L"wb");
    if (!fi || !fo) { if (fi) fclose(fi); if (fo) fclose(fo); return false; }

    mz_zip_archive r = { 0 }, w = { 0 };
    bool ok = false;
    if (mz_zip_reader_init_cfile(&r, fi, 0, 0) && mz_zip_writer_init_cfile(&w, fo, 0))
    {
        ok = true;
        mz_uint n = mz_zip_reader_get_num_files(&r);
        for (mz_uint i = 0; i < n && ok; i++)
        {
            mz_zip_archive_file_stat st;
            if (!mz_zip_reader_file_stat(&r, i, &st)) { ok = false; break; }
            if (core && strcmp(st.m_filename, "docProps/core.xml") == 0)
                ok = mz_zip_writer_add_mem(&w, st.m_filename, core->data(), core->size(), MZ_DEFAULT_COMPRESSION);
            else if (app && strcmp(st.m_filename, "docProps/app.xml") == 0)
                ok = mz_zip_writer_add_mem(&w, st.m_filename, app->data(), app->size(), MZ_DEFAULT_COMPRESSION);
            else
                ok = mz_zip_writer_add_from_zip_reader(&w, &r, i);
        }
        if (ok) ok = mz_zip_writer_finalize_archive(&w);
    }
    mz_zip_writer_end(&w);
    mz_zip_reader_end(&r);
    fclose(fi); fclose(fo);

    if (ok)
        ok = !!MoveFileExW(tmp, path, MOVEFILE_REPLACE_EXISTING);
    else
        _wremove(tmp);
    return ok;
}

bool CMainFrame::ApplyChanges()
{
    if (m_path.IsEmpty()) { Say(_T("未选择文件")); return false; }
    bool okAll = true;

    if (m_hasDoc)                                   // 无复选框: 支持即写
    {
        std::string xml = m_core;
        SYSTEMTIME loc, utc;
        bool ok = true;
        auto put = [&](CDateTimeCtrl& dtp, const char* tag) {
            if (!ok) return;
            if (dtp.GetTime(&loc) == GDT_VALID && LocalToUtc(loc, utc))
                ok = SetTag(xml, tag, UtcToIso(utc), "</cp:coreProperties>");
            else ok = false; };
        put(m_dc[0], "dcterms:created");
        put(m_dc[1], "dcterms:modified");
        put(m_dc[2], "cp:lastPrinted");
        auto putTx = [&](CEdit& e, const char* tag) {
            if (!ok) return;
            CStringW s; e.GetWindowText(s);
            ok = SetTag(xml, tag, XmlEscape(s), "</cp:coreProperties>"); };
        putTx(m_tx[0], "dc:creator");
        putTx(m_tx[1], "cp:lastModifiedBy");
        putTx(m_tx[2], "cp:revision");
        putTx(m_tx[3], "cp:version");
        std::string appXml;
        if (m_hasApp)
        {
            appXml = m_appCore;
            auto putApp = [&](CEdit& e, const char* tag) {
                if (!ok) return;
                CStringW s; e.GetWindowText(s);
                ok = SetTag(appXml, tag, XmlEscape(s), "</Properties>"); };
            putApp(m_tx[4], "Company");
            putApp(m_tx[5], "Manager");
        }
        okAll = ok && WriteCoreXml(m_path, &xml, m_hasApp ? &appXml : nullptr);
    }

    {
        HANDLE h = CreateFileW(m_path, GENERIC_WRITE, FILE_SHARE_READ, NULL,
                               OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
        if (h == INVALID_HANDLE_VALUE) { Say(_T("打开失败(可能只读/被占用): %lu"), GetLastError()); return false; }
        SYSTEMTIME loc, utc;
        FILETIME fc, fa, fw;
        auto ft = [&](CDateTimeCtrl& dtp, FILETIME& out) -> bool {
            return dtp.GetTime(&loc) == GDT_VALID && LocalToUtc(loc, utc) && !!SystemTimeToFileTime(&utc, &out); };
        bool ok = ft(m_fs[0], fc) && ft(m_fs[1], fw) && ft(m_fs[2], fa) && !!SetFileTime(h, &fc, &fa, &fw);
        CloseHandle(h);
        okAll = okAll && ok;
    }
    return okAll;
}

void CMainFrame::OnBack()
{
    if (m_page > 0) SetPage(m_page - 1);
}

void CMainFrame::OnNext()
{
    if (m_page == 0 && m_path.IsEmpty()) { Say(_T("请先选择文件")); return; }
    bool last = (m_page == 2 && m_hasDoc) || (m_page == 1 && !m_hasDoc);
    if (last)
    {
        if (m_saving) return;                    // 防重入: 大文件写入耗时长
        m_saving = true;
        m_btnNext.EnableWindow(FALSE);
        if (ApplyChanges())
        {
            Say(_T("已保存"));
            AfxMessageBox(_T("已保存。"), MB_OK | MB_ICONINFORMATION);
            AddRecent(m_path);                   // 记录本次操作
            m_saving = false;
            m_btnNext.EnableWindow(TRUE);
            m_path.Empty();                      // 回到第1屏, 重置向导
            m_core.clear();
            m_appCore.clear();
            m_hasDoc = m_hasApp = false;
            EnableDocGroup(false);
            m_editPath.SetWindowText(_T(""));
            UpdatePathTip();
            RefreshRecent();
            SetPage(0);
            Say(_T("已保存，可继续选择下一个文件"));
        }
        else
        {
            Say(_T("保存失败，请检查文件是否被占用/只读"));
            m_saving = false;
            m_btnNext.EnableWindow(TRUE);
        }
        return;
    }
    SetPage(m_page + 1);
}

class CApp : public CWinApp
{
    BOOL InitInstance() override
    {
        INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_DATE_CLASSES };
        InitCommonControlsEx(&icc);
        CMainFrame* f = new CMainFrame();
        m_pMainWnd = f;
        // 对话框样式: 不可缩放、无最大化按钮，固定尺寸
        f->Create(NULL, _T("文件时间 / Office 属性向导"),
                  WS_OVERLAPPED | WS_SYSMENU | WS_CAPTION | WS_MINIMIZEBOX,
                  CRect(0, 0, 670, 452));
        f->CenterWindow();
        f->ShowWindow(m_nCmdShow);
        f->UpdateWindow();
        // 微笑 logo: 标题栏(小) + 任务栏(大)
        HINSTANCE hi = AfxGetInstanceHandle();
        HICON hSmall = (HICON)LoadImageW(hi, MAKEINTRESOURCEW(1), IMAGE_ICON,
                                         GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR);
        HICON hBig = (HICON)LoadImageW(hi, MAKEINTRESOURCEW(1), IMAGE_ICON,
                                        GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR);
        f->SendMessage(WM_SETICON, ICON_SMALL, (LPARAM)hSmall);
        f->SendMessage(WM_SETICON, ICON_BIG, (LPARAM)hBig);
        return TRUE;
    }
} theApp;
