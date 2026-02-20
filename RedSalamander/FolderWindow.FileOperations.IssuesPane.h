#pragma once

#include "FolderWindow.h"

#include <memory>

class FileOperationsIssuesPane final
{
public:
    static HWND Create(FolderWindow::FileOperationState* fileOps,
                       FolderWindow* folderWindow,
                       HWND ownerWindow,
                       std::weak_ptr<void> hostLifetime) noexcept;

private:
    FileOperationsIssuesPane() = delete;
};
