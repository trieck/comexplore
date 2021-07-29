#pragma once

class MessageLoopEx : public CMessageLoop
{
public:
    MessageLoopEx() = default;
    virtual ~MessageLoopEx() = default;

    virtual int Run();
};
