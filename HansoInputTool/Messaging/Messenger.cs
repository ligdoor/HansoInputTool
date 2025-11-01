using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Xaml.Behaviors;

namespace HansoInputTool.Messaging
{
    public interface IMessage { }

    public static class Messenger
    {
        private static readonly Dictionary<Type, List<WeakAction>> Subscribers = new();

        public static void Register<T>(object recipient, Action<T> action) where T : IMessage
        {
            var messageType = typeof(T);
            if (!Subscribers.ContainsKey(messageType))
            {
                Subscribers[messageType] = new List<WeakAction>();
            }
            Subscribers[messageType].Add(new WeakAction<T>(recipient, action));
        }

        public static void Unregister(object recipient)
        {
            foreach (var messageType in Subscribers.Keys)
            {
                var toRemove = Subscribers[messageType].Where(wa => !wa.IsAlive || wa.Recipient == recipient).ToList();
                foreach (var weakAction in toRemove)
                {
                    Subscribers[messageType].Remove(weakAction);
                }
            }
        }

        public static void Send<T>(T message) where T : IMessage
        {
            var messageType = typeof(T);
            if (Subscribers.ContainsKey(messageType))
            {
                var deadActions = Subscribers[messageType].Where(wa => !wa.IsAlive).ToList();
                foreach (var deadAction in deadActions)
                {
                    Subscribers[messageType].Remove(deadAction);
                }
                foreach (var action in Subscribers[messageType].ToList())
                {
                    action.Execute(message);
                }
            }
        }

        private abstract class WeakAction
        {
            public abstract bool IsAlive { get; }
            public abstract object Recipient { get; }
            public abstract void Execute(IMessage message);
        }

        private class WeakAction<T> : WeakAction where T : IMessage
        {
            private readonly WeakReference _recipientRef;
            private readonly Action<T> _action;
            public override bool IsAlive => _recipientRef.IsAlive;
            public override object Recipient => _recipientRef.Target;
            public WeakAction(object recipient, Action<T> action) { _recipientRef = new WeakReference(recipient); _action = action; }
            public override void Execute(IMessage message) { if (IsAlive && Recipient != null && message is T typedMessage) { _action(typedMessage); } }
        }
    }

    public class MessageTrigger : TriggerBase<FrameworkElement>
    {
        public Type MessageType { get { return (Type)GetValue(MessageTypeProperty); } set { SetValue(MessageTypeProperty, value); } }
        public static readonly DependencyProperty MessageTypeProperty = DependencyProperty.Register("MessageType", typeof(Type), typeof(MessageTrigger), new PropertyMetadata(null));

        protected override void OnAttached()
        {
            base.OnAttached();
            if (MessageType != null)
            {
                Messenger.Register(this, (IMessage msg) => { if (msg.GetType() == MessageType) { InvokeActions(msg); } });
            }
        }
        protected override void OnDetaching()
        {
            Messenger.Unregister(this);
            base.OnDetaching();
        }
    }

    public class FocusAction : TriggerAction<FrameworkElement>
    {
        protected override void Invoke(object parameter)
        {
            if (parameter is FocusMessage message && !string.IsNullOrEmpty(message.TargetElementName))
            {
                if (AssociatedObject.FindName(message.TargetElementName) is UIElement targetElement)
                {
                    targetElement.Focus();
                }
            }
        }
    }
}